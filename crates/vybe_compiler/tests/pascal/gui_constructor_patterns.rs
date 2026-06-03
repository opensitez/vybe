//! Constructor patterns: Create, inherited, override, Setup methods.

use super::helpers::run_pascal_gui;

#[test]
fn create_constructor_creates_controls() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  btn1: TButton;
  txtName: TEdit;
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.Left := 10; btn1.Top := 20; btn1.Width := 80; btn1.Height := 30; btn1.Caption := 'Click'; txtName := TEdit.Create(Self); txtName.Name := 'txtName'; txtName.Left := 10; txtName.Top := 60; txtName.Width := 200; txtName.Height := 25; Self.Controls.Add(btn1); Self.Controls.Add(txtName); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 2);
    assert!(g.control_names.contains(&"btn1".to_string()));
    assert!(g.control_names.contains(&"txtname".to_string()));
}

#[test]
fn create_sets_form_caption() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); Caption := 'Login Form'; end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let mut g = gui.lock().unwrap();
    assert_eq!(g.get_property("form1", "Text"), "Login Form");
}

#[test]
fn create_with_event_handlers() {
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
    assert_eq!(g.control_names.len(), 1);
    assert!(g.event_handlers.contains_key("btn1.click"));
}

#[test]
fn create_with_enabled_visible() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  btn1: TButton;
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.Enabled := False; btn1.Visible := True; Self.Controls.Add(btn1); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn setup_method_pattern() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  btn1: TButton;
  constructor Create(AOwner: TObject); override;
  procedure SetupControls;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); SetupControls; end;
procedure TForm1.SetupControls;
begin btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.Caption := 'Setup'; Self.Controls.Add(btn1); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn inherited_constructor_calls_parent() {
    let out = super::helpers::run_pascal(
        r#"
program Test;
type TBase = class
  FVal: Integer;
  constructor Create;
end;
type TDerived = class(TBase)
  constructor Create;
end;
constructor TBase.Create; begin FVal := 42; end;
constructor TDerived.Create; begin inherited Create; end;
var d: TDerived;
begin d := TDerived.Create; WriteLn(d.FVal); end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn override_constructor_with_params() {
    let out = super::helpers::run_pascal(
        r#"
program Test;
type TBase = class
  FVal: Integer;
  constructor Create(V: Integer);
end;
type TDerived = class(TBase)
  constructor Create(V: Integer); override;
end;
constructor TBase.Create(V: Integer); begin FVal := V; end;
constructor TDerived.Create(V: Integer); begin inherited Create(V * 2); end;
var d: TDerived;
begin d := TDerived.Create(5); WriteLn(d.FVal); end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}
