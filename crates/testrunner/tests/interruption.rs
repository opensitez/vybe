#![cfg(unix)]

use std::process::{Command, Stdio};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use wait_timeout::ChildExt;

#[test]
fn terminating_runner_stops_its_worker_group_and_preserves_active_journal() {
    interruption_case(false);
}

#[test]
fn durable_journal_mode_preserves_active_start_on_interruption() {
    interruption_case(true);
}

fn interruption_case(durable: bool) {
    let directory = std::env::temp_dir().join(format!("testrunner-interrupt-{}-{durable}-{}",
        std::process::id(), SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_nanos()));
    std::fs::create_dir(&directory).unwrap();
    let test = directory.join("probe.js");
    std::fs::write(&test, "// vybe-test: js/probe/sleep\n").unwrap();
    let helper = directory.join("runtime.sh");
    std::fs::write(&helper, "echo $$ > \"$PID_FILE\"\nsleep 20\n").unwrap();
    let results = directory.join("results");
    let pid_file = directory.join("child.pid");
    let mut command = Command::new(env!("CARGO_BIN_EXE_testrunner"));
    command.args(["run", test.to_str().unwrap(), "--runtime",
            &format!("/bin/sh {}", helper.display()), "--cold", "-j", "1",
            "--timeout", "30", "--save", "--results", results.to_str().unwrap()])
        .env("PID_FILE", &pid_file).stdout(Stdio::null()).stderr(Stdio::null());
    if durable { command.arg("--durable-journal"); }
    let mut runner = command.spawn().unwrap();
    let started = Instant::now();
    let mut worker = None;
    // File creation precedes its contents; interruption must also follow journaling.
    while started.elapsed() < Duration::from_secs(3) {
        worker = std::fs::read_to_string(&pid_file).ok()
            .and_then(|value| value.trim().parse::<i32>().ok());
        let journal_started = std::fs::read_dir(&results).ok().into_iter().flatten()
            .filter_map(Result::ok)
            .filter(|entry| entry.file_name().to_string_lossy().starts_with("active-"))
            .filter_map(|entry| std::fs::read_to_string(entry.path()).ok())
            .flat_map(|text| text.lines().map(str::to_owned).collect::<Vec<_>>())
            .filter_map(|line| serde_json::from_str::<serde_json::Value>(&line).ok())
            .any(|event| event["event"] == "start"
                && worker.is_some_and(|pid| event["pid"] == pid));
        if worker.is_some() && journal_started { break; }
        std::thread::sleep(Duration::from_millis(10));
    }
    unsafe { libc::kill(runner.id() as i32, libc::SIGTERM); }
    let status = runner.wait_timeout(Duration::from_secs(3)).unwrap();
    if status.is_none() { let _ = runner.kill(); let _ = runner.wait(); }
    let mut stopped = false;
    if let Some(pid) = worker {
        for _ in 0..100 {
            if unsafe { libc::kill(-pid, 0) } == -1 { stopped = true; break; }
            std::thread::sleep(Duration::from_millis(10));
        }
        if !stopped { unsafe { libc::kill(-pid, libc::SIGKILL); } }
    }
    let journal = std::fs::read_dir(&results).unwrap().filter_map(Result::ok)
        .find(|entry| entry.file_name().to_string_lossy().starts_with("active-"))
        .map(|entry| std::fs::read_to_string(entry.path()).unwrap()).unwrap();
    std::fs::remove_dir_all(&directory).unwrap();
    assert!(worker.is_some(), "fake runtime failed to start");
    assert_eq!(status.unwrap().code(), Some(128 + libc::SIGTERM));
    assert!(stopped, "runner interruption left worker group alive");
    let events: Vec<serde_json::Value> = journal.lines()
        .map(|line| serde_json::from_str(line).unwrap()).collect();
    assert_eq!(events.len(), 1);
    assert_eq!(events[0]["event"], "start");
    assert_eq!(events[0]["pid"], worker.unwrap());
}
