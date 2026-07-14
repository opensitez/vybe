unit MainForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, ExtCtrls;

type
  TKanbanForm = class(TForm)
  private
    FInput: TEdit;
    FAddButton: TButton;
    FToDoList: TListBox;
    FDoingList: TListBox;
    FDoneList: TListBox;
    FMoveRightButton: TButton;
    FMoveLeftButton: TButton;
    FDeleteButton: TButton;
    procedure BuildUi;
    procedure AddTaskClick(Sender: TObject);
    procedure MoveRightClick(Sender: TObject);
    procedure MoveLeftClick(Sender: TObject);
    procedure DeleteClick(Sender: TObject);
    function ActiveList: TListBox;
  public
    constructor Create(AOwner: TComponent); override;
  end;

var
  KanbanForm: TKanbanForm;

implementation

{$R *.dfm}

constructor TKanbanForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  BuildUi;
end;

procedure TKanbanForm.BuildUi;
var
  Lbl: TLabel;
begin
  Caption := 'Kanban Lite';
  Width := 940;
  Height := 500;
  Position := poScreenCenter;

  FInput := TEdit.Create(Self);
  FInput.Parent := Self;
  FInput.Left := 20;
  FInput.Top := 20;
  FInput.Width := 640;
  FInput.Text := '';

  FAddButton := TButton.Create(Self);
  FAddButton.Parent := Self;
  FAddButton.Left := 670;
  FAddButton.Top := 18;
  FAddButton.Width := 110;
  FAddButton.Height := 28;
  FAddButton.Caption := 'Add';
  FAddButton.OnClick := AddTaskClick;

  Lbl := TLabel.Create(Self);
  Lbl.Parent := Self;
  Lbl.Left := 20;
  Lbl.Top := 60;
  Lbl.Caption := 'To Do';

  Lbl := TLabel.Create(Self);
  Lbl.Parent := Self;
  Lbl.Left := 320;
  Lbl.Top := 60;
  Lbl.Caption := 'Doing';

  Lbl := TLabel.Create(Self);
  Lbl.Parent := Self;
  Lbl.Left := 620;
  Lbl.Top := 60;
  Lbl.Caption := 'Done';

  FToDoList := TListBox.Create(Self);
  FToDoList.Parent := Self;
  FToDoList.Left := 20;
  FToDoList.Top := 80;
  FToDoList.Width := 260;
  FToDoList.Height := 340;

  FDoingList := TListBox.Create(Self);
  FDoingList.Parent := Self;
  FDoingList.Left := 320;
  FDoingList.Top := 80;
  FDoingList.Width := 260;
  FDoingList.Height := 340;

  FDoneList := TListBox.Create(Self);
  FDoneList.Parent := Self;
  FDoneList.Left := 620;
  FDoneList.Top := 80;
  FDoneList.Width := 260;
  FDoneList.Height := 340;

  FMoveRightButton := TButton.Create(Self);
  FMoveRightButton.Parent := Self;
  FMoveRightButton.Left := 20;
  FMoveRightButton.Top := 432;
  FMoveRightButton.Width := 110;
  FMoveRightButton.Height := 30;
  FMoveRightButton.Caption := 'Move Right';
  FMoveRightButton.OnClick := MoveRightClick;

  FMoveLeftButton := TButton.Create(Self);
  FMoveLeftButton.Parent := Self;
  FMoveLeftButton.Left := 140;
  FMoveLeftButton.Top := 432;
  FMoveLeftButton.Width := 110;
  FMoveLeftButton.Height := 30;
  FMoveLeftButton.Caption := 'Move Left';
  FMoveLeftButton.OnClick := MoveLeftClick;

  FDeleteButton := TButton.Create(Self);
  FDeleteButton.Parent := Self;
  FDeleteButton.Left := 760;
  FDeleteButton.Top := 432;
  FDeleteButton.Width := 120;
  FDeleteButton.Height := 30;
  FDeleteButton.Caption := 'Delete Task';
  FDeleteButton.OnClick := DeleteClick;

  FToDoList.Items.Add('Create sprint plan');
  FToDoList.Items.Add('Sketch login flow');
  FDoingList.Items.Add('Implement API client');
  FDoneList.Items.Add('Set up project board');
end;

procedure TKanbanForm.AddTaskClick(Sender: TObject);
begin
  if Trim(FInput.Text) = '' then
    Exit;

  FToDoList.Items.Add(Trim(FInput.Text));
  FInput.Clear;
  FInput.SetFocus;
end;

function TKanbanForm.ActiveList: TListBox;
begin
  if FToDoList.Focused and (FToDoList.ItemIndex >= 0) then
    Exit(FToDoList);

  if FDoingList.Focused and (FDoingList.ItemIndex >= 0) then
    Exit(FDoingList);

  if FDoneList.Focused and (FDoneList.ItemIndex >= 0) then
    Exit(FDoneList);

  if FToDoList.ItemIndex >= 0 then
    Exit(FToDoList);

  if FDoingList.ItemIndex >= 0 then
    Exit(FDoingList);

  if FDoneList.ItemIndex >= 0 then
    Exit(FDoneList);

  Result := nil;
end;

procedure TKanbanForm.MoveRightClick(Sender: TObject);
var
  Src, Dst: TListBox;
  Txt: string;
begin
  Src := ActiveList;
  if Src = nil then
    Exit;

  if Src = FToDoList then
    Dst := FDoingList
  else if Src = FDoingList then
    Dst := FDoneList
  else
    Exit;

  Txt := Src.Items[Src.ItemIndex];
  Src.Items.Delete(Src.ItemIndex);
  Dst.Items.Add(Txt);
end;

procedure TKanbanForm.MoveLeftClick(Sender: TObject);
var
  Src, Dst: TListBox;
  Txt: string;
begin
  Src := ActiveList;
  if Src = nil then
    Exit;

  if Src = FDoneList then
    Dst := FDoingList
  else if Src = FDoingList then
    Dst := FToDoList
  else
    Exit;

  Txt := Src.Items[Src.ItemIndex];
  Src.Items.Delete(Src.ItemIndex);
  Dst.Items.Add(Txt);
end;

procedure TKanbanForm.DeleteClick(Sender: TObject);
var
  Src: TListBox;
begin
  Src := ActiveList;
  if Src = nil then
    Exit;

  Src.Items.Delete(Src.ItemIndex);
end;

end.
