#[test]
fn executable_registers_only_selected_languages() {
    let version = std::process::Command::new(env!("CARGO_BIN_EXE_vx"))
        .arg("--version")
        .output()
        .expect("run vx --version");
    assert!(version.status.success());
    assert!(
        String::from_utf8(version.stdout)
            .unwrap()
            .starts_with("vx ")
    );
    let output = std::process::Command::new(env!("CARGO_BIN_EXE_vx"))
        .arg("--list-languages")
        .output()
        .expect("run vx");
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let output = String::from_utf8(output.stdout).unwrap();
    let mut names: Vec<_> = output
        .lines()
        .skip(1)
        .filter_map(|line| line.split_whitespace().next())
        .collect();
    names.sort_unstable();
    assert_eq!(names, ["c", "go", "js", "php", "python"], "{output}");
}

#[test]
fn php_output_keeps_direct_helper_calls() {
    // A helper call uses a placeholder import operand plus a direct-call
    // receipt. VM host-call fusion must honor that receipt even when the
    // placeholder happens to name a fusible host import.
    for (source, expected) in [
        ("<?php echo 42;", "42"),
        ("<?php echo 'hello';", "hello"),
        ("<?php print(42);", "42"),
        ("<?php $x=42; echo $x;", "42"),
    ] {
        let output = std::process::Command::new(env!("CARGO_BIN_EXE_vx"))
            .args(["--lang", "php", "--eval", source])
            .output()
            .expect("run PHP output case");
        assert!(
            output.status.success(),
            "{source}: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        assert_eq!(String::from_utf8(output.stdout).unwrap(), expected, "{source}");
    }
}
