//! Extended controls: TPanel, TImage, TShape, TBevel, TSplitter, TPageControl,
//! TTabSheet, TTimer (non-visual).

use super::helpers::run_pascal_gui;

#[test]
fn panel_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var pnl: TPanel;
begin inherited Create(AOwner); pnl := TPanel.Create(Self); pnl.Name := 'pnlMain'; Self.Controls.Add(pnl); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pnlmain".to_string()));
}

#[test]
fn image_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var img: TImage;
begin inherited Create(AOwner); img := TImage.Create(Self); img.Name := 'imgLogo'; Self.Controls.Add(img); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"imglogo".to_string()));
}

#[test]
fn shape_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var shp: TShape;
begin inherited Create(AOwner); shp := TShape.Create(Self); shp.Name := 'shpRect'; shp.Shape := stRectangle; Self.Controls.Add(shp); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"shprect".to_string()));
}

#[test]
fn bevel_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var bvl: TBevel;
begin inherited Create(AOwner); bvl := TBevel.Create(Self); bvl.Name := 'bvlSep'; Self.Controls.Add(bvl); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"bvlsep".to_string()));
}

#[test]
fn pagecontrol_tabsheet_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ComCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var pc: TPageControl; ts: TTabSheet;
begin inherited Create(AOwner); pc := TPageControl.Create(Self); pc.Name := 'pcMain'; ts := TTabSheet.Create(Self); ts.Name := 'tsGeneral'; ts.Caption := 'General'; ts.PageControl := pc; Self.Controls.Add(pc); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pcmain".to_string()));
    assert!(g.control_names.contains(&"tsgeneral".to_string()));
}

#[test]
fn timer_create_nonvisual() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var tmr: TTimer;
begin inherited Create(AOwner); tmr := TTimer.Create(Self); tmr.Name := 'tmrUpdate'; tmr.Interval := 1000; tmr.Enabled := True; end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"tmrupdate".to_string()));
}

#[test]
fn panel_with_nested_button() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, StdCtrls, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var pnl: TPanel; btn: TButton;
begin inherited Create(AOwner); pnl := TPanel.Create(Self); pnl.Name := 'pnl1'; btn := TButton.Create(Self); btn.Name := 'btn1'; pnl.Controls.Add(btn); Self.Controls.Add(pnl); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"pnl1".to_string()));
    assert!(g.control_names.contains(&"btn1".to_string()));
}

#[test]
fn splitter_create() {
    let (_vm, gui, _) = run_pascal_gui(r#"
program Test;
uses Forms, ExtCtrls;
type TForm1 = class(TForm)
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
var spl: TSplitter;
begin inherited Create(AOwner); spl := TSplitter.Create(Self); spl.Name := 'splMain'; Self.Controls.Add(spl); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#);
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"splmain".to_string()));
}
