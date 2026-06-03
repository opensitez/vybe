//! Menu tests: TMainMenu, TPopupMenu, TMenuItem, submenus.

use super::helpers::run_pascal_gui;

#[test]
fn mainmenu_with_items() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
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
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"mainmenu1".to_string()));
    assert!(g.control_names.contains(&"filemenu".to_string()));
    assert!(g.control_names.contains(&"editmenu".to_string()));
}

#[test]
fn menuitem_with_submenu() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, Menus;
type TForm1 = class(TForm)
  MainMenu1: TMainMenu;
  FileMenu, NewItem, OpenItem: TMenuItem;
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); MainMenu1 := TMainMenu.Create(Self); MainMenu1.Name := 'MainMenu1'; FileMenu := TMenuItem.Create(Self); FileMenu.Name := 'FileMenu'; FileMenu.Caption := 'File'; NewItem := TMenuItem.Create(Self); NewItem.Name := 'NewItem'; NewItem.Caption := 'New'; FileMenu.Add(NewItem); OpenItem := TMenuItem.Create(Self); OpenItem.Name := 'OpenItem'; OpenItem.Caption := 'Open'; FileMenu.Add(OpenItem); MainMenu1.Items.Add(FileMenu); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"newitem".to_string()));
    assert!(g.control_names.contains(&"openitem".to_string()));
}

#[test]
fn popupmenu_create() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, Menus;
type TForm1 = class(TForm)
  PopupMenu1: TPopupMenu;
  CutItem, CopyItem, PasteItem: TMenuItem;
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); PopupMenu1 := TPopupMenu.Create(Self); PopupMenu1.Name := 'PopupMenu1'; CutItem := TMenuItem.Create(Self); CutItem.Name := 'CutItem'; CutItem.Caption := 'Cut'; PopupMenu1.Items.Add(CutItem); CopyItem := TMenuItem.Create(Self); CopyItem.Name := 'CopyItem'; CopyItem.Caption := 'Copy'; PopupMenu1.Items.Add(CopyItem); PasteItem := TMenuItem.Create(Self); PasteItem.Name := 'PasteItem'; PasteItem.Caption := 'Paste'; PopupMenu1.Items.Add(PasteItem); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"popupmenu1".to_string()));
    assert!(g.control_names.contains(&"cutitem".to_string()));
    assert!(g.control_names.contains(&"copyitem".to_string()));
    assert!(g.control_names.contains(&"pasteitem".to_string()));
}

#[test]
fn menuitem_shortcut() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, Menus;
type TForm1 = class(TForm)
  MainMenu1: TMainMenu;
  SaveItem: TMenuItem;
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); MainMenu1 := TMainMenu.Create(Self); MainMenu1.Name := 'MainMenu1'; SaveItem := TMenuItem.Create(Self); SaveItem.Name := 'SaveItem'; SaveItem.Caption := 'Save'; SaveItem.ShortCut := TextToShortCut('Ctrl+S'); MainMenu1.Items.Add(SaveItem); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"saveitem".to_string()));
}

#[test]
fn menuitem_separator() {
    let (_vm, gui, _) = run_pascal_gui(
        r#"
program Test;
uses Forms, Menus;
type TForm1 = class(TForm)
  MainMenu1: TMainMenu;
  SepItem: TMenuItem;
  constructor Create(AOwner: TObject); override;
end;
constructor TForm1.Create(AOwner: TObject);
begin inherited Create(AOwner); MainMenu1 := TMainMenu.Create(Self); MainMenu1.Name := 'MainMenu1'; SepItem := TMenuItem.Create(Self); SepItem.Name := 'SepItem'; SepItem.Caption := '-'; MainMenu1.Items.Add(SepItem); end;
var f: TForm1;
begin f := TForm1.Create(nil); end.
"#,
    );
    let g = gui.lock().unwrap();
    assert!(g.control_names.contains(&"sepitem".to_string()));
}
