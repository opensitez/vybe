//! Complex forms with multiple controls, menus, and realistic patterns.

use super::helpers::run_pascal_gui;

#[test]
fn login_form_pattern() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TLoginForm = class(TForm)
  lblUser: TLabel;
  txtUser: TEdit;
  btnLogin: TButton;
  constructor Create(AOwner: TObject); override;
end;
constructor TLoginForm.Create(AOwner: TObject);
begin inherited Create(AOwner); lblUser := TLabel.Create(Self); lblUser.Name := 'lblUser'; lblUser.Caption := 'Username:'; lblUser.Left := 10; lblUser.Top := 15; txtUser := TEdit.Create(Self); txtUser.Name := 'txtUser'; txtUser.Left := 100; txtUser.Top := 10; txtUser.Width := 200; txtUser.Height := 25; btnLogin := TButton.Create(Self); btnLogin.Name := 'btnLogin'; btnLogin.Caption := 'Login'; btnLogin.Left := 100; btnLogin.Top := 50; btnLogin.Width := 80; btnLogin.Height := 30; Self.Controls.Add(lblUser); Self.Controls.Add(txtUser); Self.Controls.Add(btnLogin); end;
var f: TLoginForm;
begin f := TLoginForm.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert_eq!(g.control_names.len(), 3);
    assert!(g.control_names.contains(&"lbluser".to_string()));
    assert!(g.control_names.contains(&"txtuser".to_string()));
    assert!(g.control_names.contains(&"btnlogin".to_string()));
}

#[test]
fn settings_form_with_tabs() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls, ComCtrls, ExtCtrls;
type TSettingsForm = class(TForm)
  pc: TPageControl;
  tsGeneral, tsAdvanced: TTabSheet;
  chkAutoSave: TCheckBox;
  btnOK: TButton;
  constructor Create(AOwner: TObject); override;
end;
constructor TSettingsForm.Create(AOwner: TObject);
begin inherited Create(AOwner); pc := TPageControl.Create(Self); pc.Name := 'pcSettings'; tsGeneral := TTabSheet.Create(Self); tsGeneral.Name := 'tsGeneral'; tsGeneral.Caption := 'General'; tsGeneral.PageControl := pc; tsAdvanced := TTabSheet.Create(Self); tsAdvanced.Name := 'tsAdvanced'; tsAdvanced.Caption := 'Advanced'; tsAdvanced.PageControl := pc; chkAutoSave := TCheckBox.Create(Self); chkAutoSave.Name := 'chkAutoSave'; chkAutoSave.Caption := 'Auto Save'; Self.Controls.Add(pc); Self.Controls.Add(chkAutoSave); btnOK := TButton.Create(Self); btnOK.Name := 'btnOK'; btnOK.Caption := 'OK'; Self.Controls.Add(btnOK); end;
var f: TSettingsForm;
begin f := TSettingsForm.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pcsettings".to_string()));
    assert!(g.control_names.contains(&"tsgeneral".to_string()));
    assert!(g.control_names.contains(&"tsadvanced".to_string()));
    assert!(g.control_names.contains(&"chkautosave".to_string()));
    assert!(g.control_names.contains(&"btnok".to_string()));
}

#[test]
fn data_browser_form() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, Grids, StdCtrls, ComCtrls, ExtCtrls;
type TDataForm = class(TForm)
  sg: TStringGrid;
  pnlToolbar: TPanel;
  btnAdd, btnDelete, btnRefresh: TButton;
  txtFilter: TEdit;
  constructor Create(AOwner: TObject); override;
end;
constructor TDataForm.Create(AOwner: TObject);
begin inherited Create(AOwner); sg := TStringGrid.Create(Self); sg.Name := 'sgData'; sg.ColCount := 4; sg.RowCount := 20; Self.Controls.Add(sg); pnlToolbar := TPanel.Create(Self); pnlToolbar.Name := 'pnlToolbar'; pnlToolbar.Align := alTop; pnlToolbar.Height := 40; Self.Controls.Add(pnlToolbar); btnAdd := TButton.Create(Self); btnAdd.Name := 'btnAdd'; btnAdd.Caption := 'Add'; btnAdd.Left := 5; btnAdd.Top := 5; pnlToolbar.Controls.Add(btnAdd); btnDelete := TButton.Create(Self); btnDelete.Name := 'btnDelete'; btnDelete.Caption := 'Delete'; btnDelete.Left := 90; btnDelete.Top := 5; pnlToolbar.Controls.Add(btnDelete); btnRefresh := TButton.Create(Self); btnRefresh.Name := 'btnRefresh'; btnRefresh.Caption := 'Refresh'; btnRefresh.Left := 180; btnRefresh.Top := 5; pnlToolbar.Controls.Add(btnRefresh); txtFilter := TEdit.Create(Self); txtFilter.Name := 'txtFilter'; txtFilter.Left := 280; txtFilter.Top := 5; txtFilter.Width := 150; pnlToolbar.Controls.Add(txtFilter); end;
var f: TDataForm;
begin f := TDataForm.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"sgdata".to_string()));
    assert!(g.control_names.contains(&"pnltoolbar".to_string()));
    assert!(g.control_names.contains(&"btnadd".to_string()));
    assert!(g.control_names.contains(&"btndelete".to_string()));
    assert!(g.control_names.contains(&"btnrefresh".to_string()));
    assert!(g.control_names.contains(&"txtfilter".to_string()));
}

#[test]
fn calculator_form() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TCalcForm = class(TForm)
  txtDisplay: TEdit;
  btn0, btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn9: TButton;
  btnPlus, btnMinus, btnEquals, btnClear: TButton;
  constructor Create(AOwner: TObject); override;
end;
constructor TCalcForm.Create(AOwner: TObject);
begin inherited Create(AOwner); txtDisplay := TEdit.Create(Self); txtDisplay.Name := 'txtDisplay'; txtDisplay.Left := 10; txtDisplay.Top := 10; txtDisplay.Width := 200; txtDisplay.Height := 30; Self.Controls.Add(txtDisplay); btn7 := TButton.Create(Self); btn7.Name := 'btn7'; btn7.Caption := '7'; btn7.Left := 10; btn7.Top := 50; btn7.Width := 40; btn7.Height := 40; Self.Controls.Add(btn7); btn8 := TButton.Create(Self); btn8.Name := 'btn8'; btn8.Caption := '8'; btn8.Left := 60; btn8.Top := 50; btn8.Width := 40; btn8.Height := 40; Self.Controls.Add(btn8); btn9 := TButton.Create(Self); btn9.Name := 'btn9'; btn9.Caption := '9'; btn9.Left := 110; btn9.Top := 50; btn9.Width := 40; btn9.Height := 40; Self.Controls.Add(btn9); btnPlus := TButton.Create(Self); btnPlus.Name := 'btnPlus'; btnPlus.Caption := '+'; btnPlus.Left := 170; btnPlus.Top := 50; btnPlus.Width := 40; btnPlus.Height := 40; Self.Controls.Add(btnPlus); btn4 := TButton.Create(Self); btn4.Name := 'btn4'; btn4.Caption := '4'; btn4.Left := 10; btn4.Top := 100; btn4.Width := 40; btn4.Height := 40; Self.Controls.Add(btn4); btn5 := TButton.Create(Self); btn5.Name := 'btn5'; btn5.Caption := '5'; btn5.Left := 60; btn5.Top := 100; btn5.Width := 40; btn5.Height := 40; Self.Controls.Add(btn5); btn6 := TButton.Create(Self); btn6.Name := 'btn6'; btn6.Caption := '6'; btn6.Left := 110; btn6.Top := 100; btn6.Width := 40; btn6.Height := 40; Self.Controls.Add(btn6); btnMinus := TButton.Create(Self); btnMinus.Name := 'btnMinus'; btnMinus.Caption := '-'; btnMinus.Left := 170; btnMinus.Top := 100; btnMinus.Width := 40; btnMinus.Height := 40; Self.Controls.Add(btnMinus); btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.Caption := '1'; btn1.Left := 10; btn1.Top := 150; btn1.Width := 40; btn1.Height := 40; Self.Controls.Add(btn1); btn2 := TButton.Create(Self); btn2.Name := 'btn2'; btn2.Caption := '2'; btn2.Left := 60; btn2.Top := 150; btn2.Width := 40; btn2.Height := 40; Self.Controls.Add(btn2); btn3 := TButton.Create(Self); btn3.Name := 'btn3'; btn3.Caption := '3'; btn3.Left := 110; btn3.Top := 150; btn3.Width := 40; btn3.Height := 40; Self.Controls.Add(btn3); btnEquals := TButton.Create(Self); btnEquals.Name := 'btnEquals'; btnEquals.Caption := '='; btnEquals.Left := 170; btnEquals.Top := 150; btnEquals.Width := 40; btnEquals.Height := 90; Self.Controls.Add(btnEquals); btn0 := TButton.Create(Self); btn0.Name := 'btn0'; btn0.Caption := '0'; btn0.Left := 10; btn0.Top := 200; btn0.Width := 90; btn0.Height := 40; Self.Controls.Add(btn0); btnClear := TButton.Create(Self); btnClear.Name := 'btnClear'; btnClear.Caption := 'C'; btnClear.Left := 110; btnClear.Top := 200; btnClear.Width := 40; btnClear.Height := 40; Self.Controls.Add(btnClear); end;
var f: TCalcForm;
begin f := TCalcForm.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txtdisplay".to_string()));
    assert!(g.control_names.contains(&"btn0".to_string()));
    assert!(g.control_names.contains(&"btnplus".to_string()));
    assert!(g.control_names.contains(&"btnequals".to_string()));
}

#[test]
fn form_with_mainmenu() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, Menus;
type TForm1 = class(TForm)
  MainMenu1: TMainMenu;
  FileMenu, EditMenu: TMenuItem;
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); MainMenu1 := TMainMenu.Create(Self); MainMenu1.Name := 'MainMenu1'; FileMenu := TMenuItem.Create(Self); FileMenu.Name := 'FileMenu'; FileMenu.Caption := 'File'; MainMenu1.Items.Add(FileMenu); EditMenu := TMenuItem.Create(Self); EditMenu.Name := 'EditMenu'; EditMenu.Caption := 'Edit'; MainMenu1.Items.Add(EditMenu); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"mainmenu1".to_string()));
    assert!(g.control_names.contains(&"filemenu".to_string()));
    assert!(g.control_names.contains(&"editmenu".to_string()));
}

#[test]
fn two_forms_independent_controls() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TFormA = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
type TFormB = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TFormA.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btnA'; Self.Controls.Add(btn); end;
constructor TFormB.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btnB'; Self.Controls.Add(btn); end;
var a: TFormA; b: TFormB;
begin a := TFormA.Create(nil); b := TFormB.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btna".to_string()));
    assert!(g.control_names.contains(&"btnb".to_string()));
}

#[test]
fn constructor_calls_setup_method() {
    let out = super::helpers::run_pascal(r#"
program Test;
type TForm1 = class
  FTitle: String;
  FW: Integer;
  constructor Create;
  procedure SetupDefaults;
  function GetTitle: String;
end;
constructor TForm1.Create; begin SetupDefaults; end;
procedure TForm1.SetupDefaults; begin FTitle := 'Default'; FW := 800; end;
function TForm1.GetTitle: String; begin Result := FTitle; end;
var f: TForm1;
begin f := TForm1.Create; WriteLn(f.GetTitle); WriteLn(f.FW); end.
"#);
    assert_eq!(out, vec!["Default", "800"]);
}
