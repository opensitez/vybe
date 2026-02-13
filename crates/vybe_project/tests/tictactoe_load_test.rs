use vybe_project::serialization::load_project_auto;
use std::path::PathBuf;
// use vybe_forms::Control;

#[test]
fn test_load_tictactoe() {
    let manifest_dir = std::env::var("CARGO_MANIFEST_DIR").unwrap();
    // crates/vybe_project -> parent -> parent -> project root
    let root_dir = PathBuf::from(manifest_dir).parent().unwrap().parent().unwrap().to_path_buf();
    let project_path = root_dir.join("examples/TicTacToe/TicTacToe.vbproj");

    println!("Loading project from: {:?}", project_path);
    if !project_path.exists() {
        panic!("Project file not found at {:?}", project_path);
    }

    let project = load_project_auto(&project_path).expect("Failed to load TicTacToe project");

    assert_eq!(project.name, "TicTacToe");
    // StartupObject was TicTacToe.Form1 -> Form1 or TicTacToe.Form1
    let startup = project.startup_form.as_ref().expect("No startup form");
    assert!(startup.ends_with("Form1"), "Startup form {} should end with Form1", startup);

    // Check Forms
    assert_eq!(project.forms.len(), 1, "Should have 1 form");
    let form_file = &project.forms[0];
    let form = &form_file.form;
    
    assert_eq!(form.name, "Form1");

    // Check Controls
    // 9 buttons (btn0-8) + 1 label (lblStatus) + 1 button (btnReset) = 11 controls
    assert_eq!(form.controls.len(), 11, "Should have 11 controls");
    
    let btn0 = form.controls.iter().find(|c| c.name == "btn0").expect("btn0 missing");
    assert_eq!(btn0.control_type.as_str(), "Button");
    
    let lbl = form.controls.iter().find(|c| c.name == "lblStatus").expect("lblStatus missing");
    assert_eq!(lbl.control_type.as_str(), "Label");
    assert_eq!(lbl.get_text(), Some("Player X Turn"));
}
