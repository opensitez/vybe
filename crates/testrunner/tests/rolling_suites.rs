#![cfg(unix)]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use wait_timeout::ChildExt;

struct Fixture {
    directory: PathBuf,
    runner: Option<Child>,
}

impl Drop for Fixture {
    fn drop(&mut self) {
        if let Some(child) = &mut self.runner {
            if child.try_wait().ok().flatten().is_none() {
                unsafe { libc::kill(child.id() as i32, libc::SIGTERM); }
                if child.wait_timeout(Duration::from_secs(3)).ok().flatten().is_none() {
                    let _ = child.kill();
                    let _ = child.wait();
                }
            }
        }
        let _ = std::fs::remove_dir_all(&self.directory);
    }
}

#[test]
fn ten_suites_roll_through_three_slots_while_two_suites_remain_blocked() {
    let directory = std::env::temp_dir().join(format!("testrunner-rolling-{}-{}",
        std::process::id(), SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_nanos()));
    std::fs::create_dir_all(&directory).unwrap();
    let mut fixture = Fixture { directory, runner: None };
    let helper = fixture.directory.join("runtime.sh");
    let events = fixture.directory.join("events.txt");
    let release = fixture.directory.join("release");
    let results = fixture.directory.join("results");
    let output = fixture.directory.join("stdout.txt");
    std::fs::write(&helper, r#"printf '%s\n' "$1" >> "$EVENTS"
case "$1" in
    *slow_a.js|*slow_b.js)
        while [ ! -f "$RELEASE" ]; do sleep 0.01; done ;;
esac
"#).unwrap();
    let mut suites = Vec::new();
    for index in 0..10 {
        let name = match index { 0 => "slow_a".to_owned(), 1 => "slow_b".to_owned(),
            _ => format!("case_{index}") };
        let suite = fixture.directory.join("tests/js").join(&name);
        std::fs::create_dir_all(&suite).unwrap();
        std::fs::write(suite.join(format!("{name}.js")), "// vybe-test: js/rolling/case\n").unwrap();
        suites.push(suite);
    }
    let mut command = Command::new(env!("CARGO_BIN_EXE_testrunner"));
    command.arg("run").args(&suites).args(["--runtime", &format!("/bin/sh {}", helper.display()),
        "--cold", "-j", "3", "--suites", "3", "--timeout", "5", "--save", "--results"])
        .arg(&results).env("EVENTS", &events).env("RELEASE", &release)
        .stdout(std::fs::File::create(&output).unwrap()).stderr(Stdio::null());
    fixture.runner = Some(command.spawn().unwrap());
    let deadline = Instant::now() + Duration::from_secs(4);
    let mut reached_last = false;
    while Instant::now() < deadline {
        let text = std::fs::read_to_string(&events).unwrap_or_default();
        if text.contains("slow_a.js") && text.contains("slow_b.js") && text.contains("case_9.js") {
            reached_last = true;
            break;
        }
        if fixture.runner.as_mut().unwrap().try_wait().unwrap().is_some() { break; }
        std::thread::sleep(Duration::from_millis(10));
    }
    assert!(reached_last, "later suites waited for the two blocked suites");
    assert!(!release.exists());
    assert!(fixture.runner.as_mut().unwrap().try_wait().unwrap().is_none());
    std::fs::write(&release, b"go").unwrap();
    let status = fixture.runner.as_mut().unwrap().wait_timeout(Duration::from_secs(5)).unwrap();
    assert!(status.is_some_and(|status| status.success()));
    let summary = std::fs::read_to_string(&output).unwrap();
    assert!(summary.contains("/s"), "{summary}");
    assert!(summary.contains("10 suite(s)"), "same-language targets were merged: {summary}");
    let logs: Vec<_> = std::fs::read_dir(results.join("saved")).unwrap()
        .filter_map(Result::ok).filter(|entry| entry.path().extension().is_some_and(|ext| ext == "txt"))
        .collect();
    assert_eq!(logs.len(), 10);
    for log in logs {
        let text = std::fs::read_to_string(log.path()).unwrap();
        assert!(text.contains("1 passed; 0 failed"), "{}: {text}", log.path().display());
        assert!(text.contains("suite throughput:"), "{}: {text}", log.path().display());
    }
}

#[test]
fn default_twelve_workers_exit_after_the_last_saved_result() {
    use std::os::unix::fs::PermissionsExt;
    let directory = std::env::temp_dir().join(format!("testrunner-boot-exit-{}-{}",
        std::process::id(), SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_nanos()));
    std::fs::create_dir_all(&directory).unwrap();
    let mut fixture = Fixture { directory, runner: None };
    let helper = fixture.directory.join("fake-vybex");
    let pids = fixture.directory.join("pids");
    std::fs::write(&helper, r#"#!/bin/sh
printf '%s\n' "$$" >> "$WORKER_PIDS"
printf '##vybe-ready\n'
while IFS= read -r job; do
    printf '##vybe-result\tok\n'
done
"#).unwrap();
    std::fs::set_permissions(&helper, std::fs::Permissions::from_mode(0o700)).unwrap();
    let suite = fixture.directory.join("tests/js/boot_exit");
    std::fs::create_dir_all(&suite).unwrap();
    for index in 0..24 {
        std::fs::write(suite.join(format!("case_{index}.js")), "// vybe-test: js/boot_exit/case\n").unwrap();
    }
    for _ in 0..8 {
        let mut command = Command::new(env!("CARGO_BIN_EXE_testrunner"));
        command.arg("run").arg(&suite).arg("--vybex").arg(&helper)
            .args(["--timeout", "2", "--save", "--progress", "--results"])
            .arg(fixture.directory.join("results")).env("WORKER_PIDS", &pids)
            .stdout(Stdio::null()).stderr(Stdio::null());
        fixture.runner = Some(command.spawn().unwrap());
        let status = fixture.runner.as_mut().unwrap().wait_timeout(Duration::from_secs(5)).unwrap();
        if status.is_none() {
            // Also reap a boot whose PID has not yet returned from Command::spawn.
            for pid in std::fs::read_to_string(&pids).unwrap_or_default().lines() {
                if let Ok(pid) = pid.parse::<i32>() { unsafe { libc::kill(pid, libc::SIGKILL); } }
            }
        }
        assert!(status.is_some_and(|status| status.success()), "warm pool failed to exit after completion: {status:?}");
        let spawned = std::fs::read_to_string(&pids).unwrap();
        assert_eq!(spawned.lines().count(), 14, "expected twelve active workers plus two spares");
        for pid in spawned.lines() {
            let pid = pid.parse::<i32>().unwrap();
            assert_eq!(unsafe { libc::kill(pid, 0) }, -1, "worker {pid} survived pool shutdown");
        }
        std::fs::write(&pids, "").unwrap();
    }
}
