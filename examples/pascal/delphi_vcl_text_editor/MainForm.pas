unit MainForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, Menus, Dialogs, ExtCtrls;

type
  TTextEditorForm = class(TForm)
  private
    FMemo: TMemo;
    FStatus: TLabel;
    FMainMenu: TMainMenu;
    FOpenDialog: TOpenDialog;
    FSaveDialog: TSaveDialog;
    procedure BuildUi;
    procedure BuildMenu;
    procedure UpdateStatus;
    procedure MemoChange(Sender: TObject);
    procedure NewClick(Sender: TObject);
    procedure OpenClick(Sender: TObject);
    procedure SaveClick(Sender: TObject);
    procedure SaveAsClick(Sender: TObject);
    procedure ExitClick(Sender: TObject);
    procedure FindClick(Sender: TObject);
    procedure ToUpperClick(Sender: TObject);
    procedure ToLowerClick(Sender: TObject);
  public
    constructor Create(AOwner: TComponent); override;
  end;

var
  TextEditorForm: TTextEditorForm;

implementation

{$R *.dfm}

constructor TTextEditorForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  BuildUi;
  BuildMenu;
  UpdateStatus;
end;

procedure TTextEditorForm.BuildUi;
begin
  Caption := 'Text Editor';
  Width := 780;
  Height := 540;
  Position := poScreenCenter;

  FMemo := TMemo.Create(Self);
  FMemo.Parent := Self;
  FMemo.Align := alClient;
  FMemo.ScrollBars := ssBoth;
  FMemo.WordWrap := False;
  FMemo.Lines.Text := 'Welcome to the VCL text editor sample.';
  FMemo.OnChange := MemoChange;

  FStatus := TLabel.Create(Self);
  FStatus.Parent := Self;
  FStatus.Align := alBottom;
  FStatus.Height := 24;
end;

procedure TTextEditorForm.BuildMenu;
var
  MenuFile, MenuEdit, MenuFormat: TMenuItem;
  Item: TMenuItem;
begin
  FMainMenu := TMainMenu.Create(Self);
  Self.Menu := FMainMenu;

  MenuFile := TMenuItem.Create(FMainMenu);
  MenuFile.Caption := '&File';
  FMainMenu.Items.Add(MenuFile);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := '&New';
  Item.OnClick := NewClick;
  MenuFile.Add(Item);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := '&Open...';
  Item.OnClick := OpenClick;
  MenuFile.Add(Item);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := '&Save';
  Item.OnClick := SaveClick;
  MenuFile.Add(Item);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := 'Save &As...';
  Item.OnClick := SaveAsClick;
  MenuFile.Add(Item);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := '-';
  MenuFile.Add(Item);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := 'E&xit';
  Item.OnClick := ExitClick;
  MenuFile.Add(Item);

  MenuEdit := TMenuItem.Create(FMainMenu);
  MenuEdit.Caption := '&Edit';
  FMainMenu.Items.Add(MenuEdit);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := '&Find...';
  Item.OnClick := FindClick;
  MenuEdit.Add(Item);

  MenuFormat := TMenuItem.Create(FMainMenu);
  MenuFormat.Caption := 'F&ormat';
  FMainMenu.Items.Add(MenuFormat);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := 'To &UPPER';
  Item.OnClick := ToUpperClick;
  MenuFormat.Add(Item);

  Item := TMenuItem.Create(FMainMenu);
  Item.Caption := 'To &lower';
  Item.OnClick := ToLowerClick;
  MenuFormat.Add(Item);

  FOpenDialog := TOpenDialog.Create(Self);
  FOpenDialog.Filter := 'Text files|*.txt|All files|*.*';

  FSaveDialog := TSaveDialog.Create(Self);
  FSaveDialog.Filter := 'Text files|*.txt|All files|*.*';
  FSaveDialog.DefaultExt := 'txt';
end;

procedure TTextEditorForm.UpdateStatus;
var
  Chars, Lines: Integer;
begin
  Chars := Length(FMemo.Text);
  Lines := FMemo.Lines.Count;
  FStatus.Caption := Format('  Lines: %d   Characters: %d', [Lines, Chars]);
end;

procedure TTextEditorForm.MemoChange(Sender: TObject);
begin
  UpdateStatus;
end;

procedure TTextEditorForm.NewClick(Sender: TObject);
begin
  FMemo.Clear;
end;

procedure TTextEditorForm.OpenClick(Sender: TObject);
begin
  if FOpenDialog.Execute then
  begin
    FMemo.Lines.LoadFromFile(FOpenDialog.FileName);
    Caption := 'Text Editor - ' + ExtractFileName(FOpenDialog.FileName);
    UpdateStatus;
  end;
end;

procedure TTextEditorForm.SaveClick(Sender: TObject);
begin
  if FSaveDialog.FileName = '' then
  begin
    SaveAsClick(Sender);
    Exit;
  end;

  FMemo.Lines.SaveToFile(FSaveDialog.FileName);
end;

procedure TTextEditorForm.SaveAsClick(Sender: TObject);
begin
  if FSaveDialog.Execute then
  begin
    FMemo.Lines.SaveToFile(FSaveDialog.FileName);
    Caption := 'Text Editor - ' + ExtractFileName(FSaveDialog.FileName);
  end;
end;

procedure TTextEditorForm.ExitClick(Sender: TObject);
begin
  Close;
end;

procedure TTextEditorForm.FindClick(Sender: TObject);
var
  Query: string;
  FoundPos: Integer;
begin
  Query := InputBox('Find', 'Text to find:', '');
  if Query = '' then
    Exit;

  FoundPos := Pos(LowerCase(Query), LowerCase(FMemo.Text));
  if FoundPos > 0 then
  begin
    FMemo.SetFocus;
    FMemo.SelStart := FoundPos - 1;
    FMemo.SelLength := Length(Query);
  end
  else
    ShowMessage('Text not found.');
end;

procedure TTextEditorForm.ToUpperClick(Sender: TObject);
begin
  FMemo.Text := UpperCase(FMemo.Text);
end;

procedure TTextEditorForm.ToLowerClick(Sender: TObject);
begin
  FMemo.Text := LowerCase(FMemo.Text);
end;

end.
