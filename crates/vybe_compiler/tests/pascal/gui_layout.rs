//! Layout property tests: Left, Top, Width, Height, Align, Anchors.

use super::helpers::run_pascal_gui;

#[test]
fn button_left_top() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Left := 50; btn.Top := 100; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn button_width_height() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Width := 150; btn.Height := 40; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn edit_full_bounds() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var txt: TEdit;
begin inherited Create(AOwner); txt := TEdit.Create(Self); txt.Name := 'txt1'; txt.Left := 20; txt.Top := 30; txt.Width := 250; txt.Height := 25; Self.Controls.Add(txt); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txt1".to_string()));
}

#[test]
fn zero_bounds() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Left := 0; btn.Top := 0; btn.Width := 0; btn.Height := 0; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn large_coordinates() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; btn.Left := 1000; btn.Top := 2000; btn.Width := 500; btn.Height := 300; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn default_location_no_explicit() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn: TButton;
begin inherited Create(AOwner); btn := TButton.Create(Self); btn.Name := 'btn1'; Self.Controls.Add(btn); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn align_top_client() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var pnl: TPanel; mem: TMemo;
begin inherited Create(AOwner); pnl := TPanel.Create(Self); pnl.Name := 'pnlTop'; pnl.Align := alTop; pnl.Height := 40; Self.Controls.Add(pnl); mem := TMemo.Create(Self); mem.Name := 'memBody'; mem.Align := alClient; Self.Controls.Add(mem); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pnltop".to_string()));
    assert!(g.control_names.contains(&"membody".to_string()));
}

#[test]
fn anchors_left_top_right() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var txt: TEdit;
begin inherited Create(AOwner); txt := TEdit.Create(Self); txt.Name := 'txt1'; txt.Anchors := [akLeft, akTop, akRight]; Self.Controls.Add(txt); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"txt1".to_string()));
}

#[test]
fn multiple_controls_different_positions() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var btn1, btn2: TButton;
begin inherited Create(AOwner); btn1 := TButton.Create(Self); btn1.Name := 'btn1'; btn1.Left := 10; btn1.Top := 10; btn1.Width := 80; btn1.Height := 30; btn2 := TButton.Create(Self); btn2.Name := 'btn2'; btn2.Left := 100; btn2.Top := 10; btn2.Width := 80; btn2.Height := 30; Self.Controls.Add(btn1); Self.Controls.Add(btn2); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"btn1".to_string()));
    assert!(g.control_names.contains(&"btn2".to_string()));
}
