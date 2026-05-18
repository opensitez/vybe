//! Form creation, lifecycle, and application-level GUI tests.

use super::helpers::run_pascal_gui;

#[test]
fn form_create_emits_text_property() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let mut g = gui.lock().unwrap();
    let text = g.get_property("form1", "text");
    assert!(!text.is_empty(), "Expected a Text property for form1");
}

#[test]
fn application_run_sets_should_run() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); Application.Run; end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.should_run, "Expected should_run after Application.Run");
}

#[test]
fn form_show_sets_should_run() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); f.Show; end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.should_run, "Expected should_run after Form.Show");
}

#[test]
fn form_close_sets_close_requested() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); f.Close; end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.close_requested, "Expected close_requested after Form.Close");
}

#[test]
fn form_caption_assignment_mirrors() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); Caption := 'My Application'; end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let mut g = gui.lock().unwrap();
    let text = g.get_property("form1", "Text");
    assert_eq!(text, "My Application");
}

#[test]
fn multiple_forms_created() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
type TForm2 = class(TForm) end;
var f1: TForm1; f2: TForm2;
begin f1 := TForm1.Create(nil); f2 := TForm2.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.len() >= 2, "Expected at least 2 forms");
}

#[test]
fn empty_form_class_no_constructor() {
    let (_vm, _gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
}

#[test]
fn form_inherits_tform_explicitly() {
    let (_vm, _gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
}

#[test]
fn form_showmodal_sets_should_run() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); f.ShowModal; end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.should_run, "Expected should_run after ShowModal");
}

#[test]
fn form_client_size_settings() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); ClientWidth := 1024; ClientHeight := 768; end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.width >= 1024 || g.height >= 768, "Expected form size to be set");
}
