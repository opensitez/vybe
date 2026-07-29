//! Event handling tests: OnClick, OnChange, OnCreate, OnTimer, OnClose.

use super::helpers::run_pascal_gui;

#[test]
fn button_onclick_registered() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  btn1: TButton;
  constructor Create(AOwner: TObject); override;
  procedure btn1Click(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.OnClick := btn1Click; Self.Controls.Add(btn1); end;
procedure TForm1.btn1Click(Sender: TObject); begin WriteLn('clicked'); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(
        g.event_handlers.contains_key("btn1.click"),
        "Expected Click handler, got keys: {:?}",
        g.event_keys()
    );
}

#[test]
fn two_buttons_different_onclick() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  btn1, btn2: TButton;
  constructor Create(AOwner: TObject); override;
  procedure btn1Click(Sender: TObject);
  procedure btn2Click(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.OnClick := btn1Click; btn2 := TButton.Create(Self); btn2.Name := 'btn2'; btn2.OnClick := btn2Click; Self.Controls.Add(btn1); Self.Controls.Add(btn2); end;
procedure TForm1.btn1Click(Sender: TObject); begin end;
procedure TForm1.btn2Click(Sender: TObject); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("btn1.click"));
    assert!(g.event_handlers.contains_key("btn2.click"));
}

#[test]
fn edit_onchange_registered() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  txt1: TEdit;
  constructor Create(AOwner: TObject); override;
  procedure txt1Change(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); txt1 := TEdit.Create(Self); txt1.Name := 'txt1'; txt1.OnChange := txt1Change; Self.Controls.Add(txt1); end;
procedure TForm1.txt1Change(Sender: TObject); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("txt1.change"));
}

#[test]
fn checkbox_onclick_registered() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  chk1: TCheckBox;
  constructor Create(AOwner: TObject); override;
  procedure chk1Click(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); chk1 := TCheckBox.Create(Self); chk1.Name := 'chk1'; chk1.OnClick := chk1Click; Self.Controls.Add(chk1); end;
procedure TForm1.chk1Click(Sender: TObject); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("chk1.click"));
}

#[test]
fn combobox_onchange_registered() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  cbo1: TComboBox;
  constructor Create(AOwner: TObject); override;
  procedure cbo1Change(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); cbo1 := TComboBox.Create(Self); cbo1.Name := 'cbo1'; cbo1.OnChange := cbo1Change; Self.Controls.Add(cbo1); end;
procedure TForm1.cbo1Change(Sender: TObject); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("cbo1.change"));
}

#[test]
fn form_oncreate_registered() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
  procedure FormCreate(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); OnCreate := FormCreate; end;
procedure TForm1.FormCreate(Sender: TObject); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("form1.create"));
}

#[test]
fn timer_ontimer_registered() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, ExtCtrls;
type TForm1 = class(TForm)
  tmr1: TTimer;
  constructor Create(AOwner: TObject); override;
  procedure tmr1Timer(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); tmr1 := TTimer.Create(Self); tmr1.Name := 'tmr1'; tmr1.Interval := 500; tmr1.OnTimer := tmr1Timer; tmr1.Enabled := True; end;
procedure TForm1.tmr1Timer(Sender: TObject); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("tmr1.timer"));
}

#[test]
fn form_onclose_registered() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
  procedure FormClose(Sender: TObject; var Action: TCloseAction);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); OnClose := FormClose; end;
procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("form1.close"));
}

#[test]
fn handler_is_callable_value() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  btn1: TButton;
  constructor Create(AOwner: TObject); override;
  procedure btn1Click(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.OnClick := btn1Click; Self.Controls.Add(btn1); end;
procedure TForm1.btn1Click(Sender: TObject); begin end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    let handler = g.get_event_handler("btn1", "Click");
    assert!(handler.is_some());
    let h = handler.unwrap();
    assert!(
        !matches!(h, vybe_runtime::Value::Null),
        "Handler should not be Null"
    );
}

#[test]
fn handler_references_form_field() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  btn1: TButton;
  FCounter: Integer;
  constructor Create(AOwner: TObject); override;
  procedure btn1Click(Sender: TObject);
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); FCounter := 0; btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.OnClick := btn1Click; Self.Controls.Add(btn1); end;
procedure TForm1.btn1Click(Sender: TObject); begin FCounter := FCounter + 1; end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.event_handlers.contains_key("btn1.click"));
}
