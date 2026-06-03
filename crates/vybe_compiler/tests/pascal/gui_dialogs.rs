//! Dialog tests: ShowMessage, MessageDlg, InputBox, Close, ShowModal.

use super::helpers::{run_pascal_gui, run_pascal_gui_capture_msgbox};

#[test]
fn showmessage_emits_call() {
    let (_vm, _gui, msgs) = run_pascal_gui_capture_msgbox(
        r#"
program Test;
begin ShowMessage('Hello!'); end.
"#,
    );
    let msgs = msgs.lock().unwrap();
    assert!(
        msgs.iter().any(|(text, _)| text == "Hello!"),
        "Expected ShowMessage, got {:?}",
        *msgs
    );
}

#[test]
fn messagedlg_emits_call() {
    let (_vm, _gui, msgs) = run_pascal_gui_capture_msgbox(
        r#"
program Test;
begin MessageDlg('Are you sure?', mtConfirmation, [mbYes, mbNo], 0); end.
"#,
    );
    let msgs = msgs.lock().unwrap();
    assert!(
        msgs.iter().any(|(text, _)| text == "Are you sure?"),
        "Expected MessageDlg, got {:?}",
        *msgs
    );
}

#[test]
fn showmessage_from_method() {
    let (_vm, _gui, msgs) = run_pascal_gui_capture_msgbox(
        r#"
program Test;
type TForm1 = class
  constructor Create;
  procedure ShowMsg;
end;
constructor TForm1.Create; begin ShowMsg; end;
procedure TForm1.ShowMsg; begin ShowMessage('From method'); end;
var f: TForm1;
begin f := TForm1.Create; end.
"#,
    );
    let msgs = msgs.lock().unwrap();
    assert!(msgs.iter().any(|(text, _)| text == "From method"));
}

#[test]
fn form_close_from_method() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
  procedure CloseForm;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); CloseForm; end;
procedure TForm1.CloseForm; begin Self.Close; end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.close_requested, "Expected close_requested");
}

#[test]
fn application_run_sets_flag() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); Application.Run; end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.should_run);
}

#[test]
fn showmodal_sets_flag() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms;
type TForm1 = class(TForm) end;
var f: TForm1;
begin f := TForm1.Create(nil); f.ShowModal; end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.should_run);
}

#[test]
fn multiple_showmessages() {
    let (_vm, _gui, msgs) = run_pascal_gui_capture_msgbox(
        r#"
program Test;
begin ShowMessage('First'); ShowMessage('Second'); ShowMessage('Third'); end.
"#,
    );
    let msgs = msgs.lock().unwrap();
    assert_eq!(msgs.len(), 3);
    assert!(msgs.iter().any(|(t, _)| t == "First"));
    assert!(msgs.iter().any(|(t, _)| t == "Second"));
    assert!(msgs.iter().any(|(t, _)| t == "Third"));
}

#[test]
fn showmessage_empty_string() {
    let (_vm, _gui, msgs) = run_pascal_gui_capture_msgbox(
        r#"
program Test;
begin ShowMessage(''); end.
"#,
    );
    let msgs = msgs.lock().unwrap();
    assert!(msgs.iter().any(|(text, _)| text.is_empty()));
}
