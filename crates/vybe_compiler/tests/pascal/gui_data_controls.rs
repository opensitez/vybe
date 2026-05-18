//! Data-oriented controls: TStringGrid, TListView, TTreeView, TTrackBar,
//! TProgressBar, TSpinEdit, TColorDialog.

use super::helpers::run_pascal_gui;

#[test]
fn stringgrid_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, Grids;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var sg: TStringGrid;
begin inherited Create(AOwner); sg := TStringGrid.Create(Self); sg.Name := 'sgData'; sg.ColCount := 5; sg.RowCount := 10; Self.Controls.Add(sg); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"sgdata".to_string()));
}

#[test]
fn trackbar_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ComCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var tb: TTrackBar;
begin inherited Create(AOwner); tb := TTrackBar.Create(Self); tb.Name := 'tbVolume'; tb.Min := 0; tb.Max := 100; tb.Position := 50; Self.Controls.Add(tb); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"tbvolume".to_string()));
}

#[test]
fn progressbar_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ComCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var pb: TProgressBar;
begin inherited Create(AOwner); pb := TProgressBar.Create(Self); pb.Name := 'pbStatus'; pb.Min := 0; pb.Max := 100; pb.Position := 33; Self.Controls.Add(pb); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pbstatus".to_string()));
}

#[test]
fn spinedit_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var se: TSpinEdit;
begin inherited Create(AOwner); se := TSpinEdit.Create(Self); se.Name := 'seCount'; se.MinValue := 0; se.MaxValue := 100; se.Value := 5; Self.Controls.Add(se); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"secount".to_string()));
}

#[test]
fn listview_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ComCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var lv: TListView;
begin inherited Create(AOwner); lv := TListView.Create(Self); lv.Name := 'lvFiles'; lv.ViewStyle := vsReport; Self.Controls.Add(lv); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"lvfiles".to_string()));
}

#[test]
fn treeview_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ComCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var tv: TTreeView;
begin inherited Create(AOwner); tv := TTreeView.Create(Self); tv.Name := 'tvNodes'; Self.Controls.Add(tv); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"tvnodes".to_string()));
}

#[test]
fn colordialog_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, Dialogs;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var dlg: TColorDialog;
begin inherited Create(AOwner); dlg := TColorDialog.Create(Self); dlg.Name := 'dlgColor'; dlg.Color := clRed; end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"dlgcolor".to_string()));
}

#[test]
fn stringgrid_cell_assignment() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, Grids;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var sg: TStringGrid;
begin inherited Create(AOwner); sg := TStringGrid.Create(Self); sg.Name := 'sg1'; sg.Cells[0, 0] := 'Header'; sg.Cells[1, 1] := 'Data'; Self.Controls.Add(sg); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"sg1".to_string()));
}
