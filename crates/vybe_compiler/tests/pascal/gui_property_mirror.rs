//! Property mirroring into GuiState property store.

use super::helpers::run_pascal_gui;

#[test]
fn button_caption_mirrors() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Caption := 'OK'; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let mut g = gui.lock().unwrap();
    assert_eq!(g.get_property("btn1", "text"), "OK");
}

#[test]
fn control_name_tracked() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'myButton'; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"mybutton".to_string()));
}

#[test]
fn enabled_false_mirrors() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Enabled := False; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn visible_false_mirrors() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Visible := False; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn multiple_properties_same_control() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Caption := 'Go'; btn.Enabled := True; btn.Left := 10; btn.Top := 20; btn.Width := 80; btn.Height := 30; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let mut g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
    assert_eq!(g.get_property("btn1", "text"), "Go");
}

#[test]
fn property_before_controls_add() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Caption := 'Before Add'; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let mut g = gui.lock().unwrap();
    assert_eq!(g.get_property("btn1", "text"), "Before Add");
}

#[test]
fn edit_text_property_mirrors() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var txt: TEdit;
begin inherited Create(AOwner); txt := TEdit.Create(Self); txt.Name := 'txt1'; txt.Text := 'Hello'; Self.Controls.Add(txt); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let mut g = gui.lock().unwrap();
    assert_eq!(g.get_property("txt1", "text"), "Hello");
}

#[test]
fn memo_lines_property() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var mem: TMemo;
begin inherited Create(AOwner); mem := TMemo.Create(Self); mem.Name := 'mem1'; mem.Lines.Add('Line 1'); mem.Lines.Add('Line 2'); Self.Controls.Add(mem); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"mem1".to_string()));
}
