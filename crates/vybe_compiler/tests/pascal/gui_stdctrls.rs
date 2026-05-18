//! Standard controls: TButton, TLabel, TEdit, TCheckBox, TRadioButton,
//! TComboBox, TListBox, TMemo, TGroupBox.

use super::helpers::run_pascal_gui;

#[test]
fn button_create_and_name() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btnOK'; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btnok".to_string()));
}

#[test]
fn edit_create_with_bounds() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var txt: TEdit;
begin inherited Create(AOwner); txt := TEdit.Create(Self); txt.Name := 'txtInput'; txt.Left := 10; txt.Top := 20; txt.Width := 200; txt.Height := 25; Self.Controls.Add(txt); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txtinput".to_string()));
}

#[test]
fn label_create_with_caption() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var lbl: TLabel;
begin inherited Create(AOwner); lbl := TLabel.Create(Self); lbl.Name := 'lblTitle'; lbl.Caption := 'Welcome'; Self.Controls.Add(lbl); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let mut g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"lbltitle".to_string()));
    assert_eq!(g.get_property("lbltitle", "text"), "Welcome");
}

#[test]
fn checkbox_create_with_caption() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var chk: TCheckBox;
begin inherited Create(AOwner); chk := TCheckBox.Create(Self); chk.Name := 'chkAgree'; chk.Caption := 'I agree'; Self.Controls.Add(chk); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"chkagree".to_string()));
}

#[test]
fn radiobutton_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var rb: TRadioButton;
begin inherited Create(AOwner); rb := TRadioButton.Create(Self); rb.Name := 'rbOption1'; rb.Caption := 'Option 1'; Self.Controls.Add(rb); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"rboption1".to_string()));
}

#[test]
fn combobox_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var cbo: TComboBox;
begin inherited Create(AOwner); cbo := TComboBox.Create(Self); cbo.Name := 'cboItems'; Self.Controls.Add(cbo); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"cboitems".to_string()));
}

#[test]
fn listbox_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var lst: TListBox;
begin inherited Create(AOwner); lst := TListBox.Create(Self); lst.Name := 'lstItems'; Self.Controls.Add(lst); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"lstitems".to_string()));
}

#[test]
fn memo_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var mem: TMemo;
begin inherited Create(AOwner); mem := TMemo.Create(Self); mem.Name := 'memNotes'; Self.Controls.Add(mem); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"memnotes".to_string()));
}

#[test]
fn groupbox_create_with_caption() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var grp: TGroupBox;
begin inherited Create(AOwner); grp := TGroupBox.Create(Self); grp.Name := 'grpSettings'; grp.Caption := 'Settings'; Self.Controls.Add(grp); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"grpsettings".to_string()));
}

#[test]
fn edit_password_char() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var txt: TEdit;
begin inherited Create(AOwner); txt := TEdit.Create(Self); txt.Name := 'txtPass'; txt.PasswordChar := '*'; Self.Controls.Add(txt); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txtpass".to_string()));
}

#[test]
fn checkbox_checked_property() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var chk: TCheckBox;
begin inherited Create(AOwner); chk := TCheckBox.Create(Self); chk.Name := 'chk1'; chk.Checked := True; Self.Controls.Add(chk); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"chk1".to_string()));
}

#[test]
fn radiobutton_checked_property() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var rb: TRadioButton;
begin inherited Create(AOwner); rb := TRadioButton.Create(Self); rb.Name := 'rb1'; rb.Checked := True; Self.Controls.Add(rb); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"rb1".to_string()));
}
