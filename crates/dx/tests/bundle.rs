#[test]
fn executable_registers_only_selected_languages() {
    let version = std::process::Command::new(env!("CARGO_BIN_EXE_dx"))
        .arg("--version")
        .output()
        .expect("run dx --version");
    assert!(version.status.success());
    assert!(
        String::from_utf8(version.stdout)
            .unwrap()
            .starts_with("dx ")
    );
    let output = std::process::Command::new(env!("CARGO_BIN_EXE_dx"))
        .arg("--list-languages")
        .output()
        .expect("run dx");
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
    assert_eq!(names, ["csharp", "powershell", "vb"], "{output}");
}
