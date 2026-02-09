use irys_project::serialization::load_project_auto;
use std::path::PathBuf;

#[test]
fn test_load_vbproj() {
    let manifest_dir = std::env::var("CARGO_MANIFEST_DIR").unwrap();
    // crates/irys_project/tests/../../tests/sample_project
    // manifest_dir is crates/irys_project
    let root_dir = PathBuf::from(manifest_dir).parent().unwrap().parent().unwrap().to_path_buf();
    let project_path = root_dir.join("tests/sample_project/Sample.vbproj");

    println!("Loading project from: {:?}", project_path);
    let project = load_project_auto(&project_path).expect("Failed to load project");

    println!("Project Name: {}", project.name);
    assert_eq!(project.name, "SampleProject");
    
    println!("Startup Form: {:?}", project.startup_form);
    // StartupObject was SampleProject.Form1, normalized to Form1
    assert_eq!(project.startup_form, Some("Form1".to_string()));
    
    // Check forms
    println!("Forms: {}", project.forms.len());
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.forms[0].form.name, "Form1");

    // Check modules
    println!("Code Files: {}", project.code_files.len());
    assert_eq!(project.code_files.len(), 1);
    assert_eq!(project.code_files[0].name, "Module1");
}
