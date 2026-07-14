unit DataForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, ExtCtrls, Grids, ComCtrls, Dialogs;

type
  TPerson = record
    ID: Integer;
    Name: string;
    Email: string;
    Age: Integer;
    Active: Boolean;
  end;

  TfrmData = class(TForm)
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnSearchClick(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure sgDataSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure edtSearchChange(Sender: TObject);
    procedure chkShowActiveOnlyChange(Sender: TObject);
    procedure cmbSortByChange(Sender: TObject);
    procedure rgViewModeClick(Sender: TObject);
    procedure lstDetailsClick(Sender: TObject);
    procedure pgDetailsChange(Sender: TObject);
    procedure btnFirstClick(Sender: TObject);
    procedure btnPrevClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnLastClick(Sender: TObject);
  private
    FLabelTitle: TLabel;
    FLabelCount: TLabel;
    FLabelSelected: TLabel;
    FPanelToolbar: TPanel;
    FPanelGrid: TPanel;
    FPanelDetails: TPanel;
    FPanelNavigation: TPanel;
    FPanelFilter: TPanel;
    FSgData: TStringGrid;
    FLstDetails: TListBox;
    FEdtSearch: TEdit;
    FChkShowActiveOnly: TCheckBox;
    FCmbSortBy: TComboBox;
    FRgViewMode: TRadioGroup;
    FPgDetails: TPageControl;
    FTabGeneral: TTabSheet;
    FTabExtra: TTabSheet;
    FBtnAdd: TButton;
    FBtnEdit: TButton;
    FBtnDelete: TButton;
    FBtnSearch: TButton;
    FBtnRefresh: TButton;
    FBtnExport: TButton;
    FBtnSelect: TButton;
    FBtnClose: TButton;
    FBtnFirst: TButton;
    FBtnPrev: TButton;
    FBtnNext: TButton;
    FBtnLast: TButton;
    FPeople: array of TPerson;
    FFiltered: array of Integer;
    FCurrentIndex: Integer;
    FSelectedRecord: string;
    procedure BuildControls;
    procedure LoadSampleData;
    procedure RefreshGrid;
    procedure FilterData;
    procedure SortData;
    procedure ShowDetails(Index: Integer);
    procedure UpdateNavButtons;
    function GenerateID: Integer;
  public
    property SelectedRecord: string read FSelectedRecord;
  end;

implementation

procedure TfrmData.BuildControls;
begin
  Caption := 'Data Browser';
  Width := 780;
  Height := 580;
  Position := poMainFormCenter;

  FLabelTitle := TLabel.Create(Self);
  FLabelTitle.Parent := Self;
  FLabelTitle.Left := 20;
  FLabelTitle.Top := 10;
  FLabelTitle.Caption := 'Person Data Browser';
  FLabelTitle.Font.Size := 14;
  FLabelTitle.Font.Style := [fsBold];

  FPanelFilter := TPanel.Create(Self);
  FPanelFilter.Parent := Self;
  FPanelFilter.Left := 20;
  FPanelFilter.Top := 40;
  FPanelFilter.Width := 720;
  FPanelFilter.Height := 40;
  FPanelFilter.Caption := '';
  FPanelFilter.BevelOuter := bvLowered;

  with TLabel.Create(Self) do
  begin
    Parent := FPanelFilter;
    Left := 10;
    Top := 12;
    Caption := 'Search:';
  end;

  FEdtSearch := TEdit.Create(Self);
  FEdtSearch.Parent := FPanelFilter;
  FEdtSearch.Left := 60;
  FEdtSearch.Top := 8;
  FEdtSearch.Width := 200;
  FEdtSearch.OnChange := edtSearchChange;

  FChkShowActiveOnly := TCheckBox.Create(Self);
  FChkShowActiveOnly.Parent := FPanelFilter;
  FChkShowActiveOnly.Left := 280;
  FChkShowActiveOnly.Top := 10;
  FChkShowActiveOnly.Caption := 'Active only';
  FChkShowActiveOnly.OnClick := chkShowActiveOnlyChange;

  with TLabel.Create(Self) do
  begin
    Parent := FPanelFilter;
    Left := 400;
    Top := 12;
    Caption := 'Sort by:';
  end;

  FCmbSortBy := TComboBox.Create(Self);
  FCmbSortBy.Parent := FPanelFilter;
  FCmbSortBy.Left := 460;
  FCmbSortBy.Top := 8;
  FCmbSortBy.Width := 120;
  FCmbSortBy.Items.Add('ID');
  FCmbSortBy.Items.Add('Name');
  FCmbSortBy.Items.Add('Age');
  FCmbSortBy.ItemIndex := 0;
  FCmbSortBy.OnChange := cmbSortByChange;

  FRgViewMode := TRadioGroup.Create(Self);
  FRgViewMode.Parent := FPanelFilter;
  FRgViewMode.Left := 600;
  FRgViewMode.Top := 2;
  FRgViewMode.Width := 110;
  FRgViewMode.Height := 35;
  FRgViewMode.Caption := '';
  FRgViewMode.Items.Add('Grid');
  FRgViewMode.Items.Add('List');
  FRgViewMode.ItemIndex := 0;
  FRgViewMode.Columns := 2;
  FRgViewMode.OnClick := rgViewModeClick;

  FPanelToolbar := TPanel.Create(Self);
  FPanelToolbar.Parent := Self;
  FPanelToolbar.Left := 20;
  FPanelToolbar.Top := 90;
  FPanelToolbar.Width := 720;
  FPanelToolbar.Height := 45;
  FPanelToolbar.Caption := '';

  FBtnAdd := TButton.Create(Self);
  FBtnAdd.Parent := FPanelToolbar;
  FBtnAdd.Left := 10;
  FBtnAdd.Top := 8;
  FBtnAdd.Width := 70;
  FBtnAdd.Height := 28;
  FBtnAdd.Caption := 'Add';
  FBtnAdd.OnClick := btnAddClick;

  FBtnEdit := TButton.Create(Self);
  FBtnEdit.Parent := FPanelToolbar;
  FBtnEdit.Left := 90;
  FBtnEdit.Top := 8;
  FBtnEdit.Width := 70;
  FBtnEdit.Height := 28;
  FBtnEdit.Caption := 'Edit';
  FBtnEdit.OnClick := btnEditClick;

  FBtnDelete := TButton.Create(Self);
  FBtnDelete.Parent := FPanelToolbar;
  FBtnDelete.Left := 170;
  FBtnDelete.Top := 8;
  FBtnDelete.Width := 70;
  FBtnDelete.Height := 28;
  FBtnDelete.Caption := 'Delete';
  FBtnDelete.OnClick := btnDeleteClick;

  FBtnSearch := TButton.Create(Self);
  FBtnSearch.Parent := FPanelToolbar;
  FBtnSearch.Left := 260;
  FBtnSearch.Top := 8;
  FBtnSearch.Width := 80;
  FBtnSearch.Height := 28;
  FBtnSearch.Caption := 'Find';
  FBtnSearch.OnClick := btnSearchClick;

  FBtnRefresh := TButton.Create(Self);
  FBtnRefresh.Parent := FPanelToolbar;
  FBtnRefresh.Left := 350;
  FBtnRefresh.Top := 8;
  FBtnRefresh.Width := 80;
  FBtnRefresh.Height := 28;
  FBtnRefresh.Caption := 'Refresh';
  FBtnRefresh.OnClick := btnRefreshClick;

  FBtnExport := TButton.Create(Self);
  FBtnExport.Parent := FPanelToolbar;
  FBtnExport.Left := 440;
  FBtnExport.Top := 8;
  FBtnExport.Width := 80;
  FBtnExport.Height := 28;
  FBtnExport.Caption := 'Export';
  FBtnExport.OnClick := btnExportClick;

  FBtnSelect := TButton.Create(Self);
  FBtnSelect.Parent := FPanelToolbar;
  FBtnSelect.Left := 560;
  FBtnSelect.Top := 8;
  FBtnSelect.Width := 70;
  FBtnSelect.Height := 28;
  FBtnSelect.Caption := 'Select';
  FBtnSelect.ModalResult := mrOK;
  FBtnSelect.OnClick := btnSelectClick;

  FBtnClose := TButton.Create(Self);
  FBtnClose.Parent := FPanelToolbar;
  FBtnClose.Left := 640;
  FBtnClose.Top := 8;
  FBtnClose.Width := 70;
  FBtnClose.Height := 28;
  FBtnClose.Caption := 'Close';
  FBtnClose.ModalResult := mrCancel;
  FBtnClose.OnClick := btnCloseClick;

  FPanelGrid := TPanel.Create(Self);
  FPanelGrid.Parent := Self;
  FPanelGrid.Left := 20;
  FPanelGrid.Top := 145;
  FPanelGrid.Width := 460;
  FPanelGrid.Height := 280;
  FPanelGrid.Caption := '';

  FSgData := TStringGrid.Create(Self);
  FSgData.Parent := FPanelGrid;
  FSgData.Left := 0;
  FSgData.Top := 0;
  FSgData.Width := 460;
  FSgData.Height := 280;
  FSgData.ColCount := 5;
  FSgData.FixedCols := 0;
  FSgData.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect];
  FSgData.OnSelectCell := sgDataSelectCell;

  FSgData.Cells[0, 0] := 'ID';
  FSgData.Cells[1, 0] := 'Name';
  FSgData.Cells[2, 0] := 'Email';
  FSgData.Cells[3, 0] := 'Age';
  FSgData.Cells[4, 0] := 'Active';

  FPanelDetails := TPanel.Create(Self);
  FPanelDetails.Parent := Self;
  FPanelDetails.Left := 500;
  FPanelDetails.Top := 145;
  FPanelDetails.Width := 240;
  FPanelDetails.Height := 280;
  FPanelDetails.Caption := '';

  FPgDetails := TPageControl.Create(Self);
  FPgDetails.Parent := FPanelDetails;
  FPgDetails.Left := 0;
  FPgDetails.Top := 0;
  FPgDetails.Width := 240;
  FPgDetails.Height := 280;
  FPgDetails.OnChange := pgDetailsChange;

  FTabGeneral := TTabSheet.Create(FPgDetails);
  FTabGeneral.PageControl := FPgDetails;
  FTabGeneral.Caption := 'General';

  FLstDetails := TListBox.Create(Self);
  FLstDetails.Parent := FTabGeneral;
  FLstDetails.Left := 5;
  FLstDetails.Top := 5;
  FLstDetails.Width := 220;
  FLstDetails.Height := 240;
  FLstDetails.OnClick := lstDetailsClick;

  FTabExtra := TTabSheet.Create(FPgDetails);
  FTabExtra.PageControl := FPgDetails;
  FTabExtra.Caption := 'Stats';

  with TLabel.Create(Self) do
  begin
    Parent := FTabExtra;
    Left := 10;
    Top := 10;
    Caption := 'Record Statistics';
    Font.Style := [fsBold];
  end;

  FLabelCount := TLabel.Create(Self);
  FLabelCount.Parent := FTabExtra;
  FLabelCount.Left := 10;
  FLabelCount.Top := 40;
  FLabelCount.Caption := 'Total records: 0';

  FLabelSelected := TLabel.Create(Self);
  FLabelSelected.Parent := FTabExtra;
  FLabelSelected.Left := 10;
  FLabelSelected.Top := 65;
  FLabelSelected.Caption := 'Selected: none';

  FPanelNavigation := TPanel.Create(Self);
  FPanelNavigation.Parent := Self;
  FPanelNavigation.Left := 20;
  FPanelNavigation.Top := 440;
  FPanelNavigation.Width := 460;
  FPanelNavigation.Height := 40;
  FPanelNavigation.Caption := '';

  FBtnFirst := TButton.Create(Self);
  FBtnFirst.Parent := FPanelNavigation;
  FBtnFirst.Left := 10;
  FBtnFirst.Top := 6;
  FBtnFirst.Width := 60;
  FBtnFirst.Height := 26;
  FBtnFirst.Caption := '|<';
  FBtnFirst.OnClick := btnFirstClick;

  FBtnPrev := TButton.Create(Self);
  FBtnPrev.Parent := FPanelNavigation;
  FBtnPrev.Left := 80;
  FBtnPrev.Top := 6;
  FBtnPrev.Width := 60;
  FBtnPrev.Height := 26;
  FBtnPrev.Caption := '<';
  FBtnPrev.OnClick := btnPrevClick;

  FBtnNext := TButton.Create(Self);
  FBtnNext.Parent := FPanelNavigation;
  FBtnNext.Left := 150;
  FBtnNext.Top := 6;
  FBtnNext.Width := 60;
  FBtnNext.Height := 26;
  FBtnNext.Caption := '>';
  FBtnNext.OnClick := btnNextClick;

  FBtnLast := TButton.Create(Self);
  FBtnLast.Parent := FPanelNavigation;
  FBtnLast.Left := 220;
  FBtnLast.Top := 6;
  FBtnLast.Width := 60;
  FBtnLast.Height := 26;
  FBtnLast.Caption := '>|';
  FBtnLast.OnClick := btnLastClick;
end;

function TfrmData.GenerateID: Integer;
begin
  Result := Length(FPeople) + 1;
end;

procedure TfrmData.LoadSampleData;
var
  I: Integer;
begin
  SetLength(FPeople, 8);

  FPeople[0].ID := 1;
  FPeople[0].Name := 'Alice Johnson';
  FPeople[0].Email := 'alice@example.com';
  FPeople[0].Age := 28;
  FPeople[0].Active := True;

  FPeople[1].ID := 2;
  FPeople[1].Name := 'Bob Smith';
  FPeople[1].Email := 'bob@example.com';
  FPeople[1].Age := 34;
  FPeople[1].Active := True;

  FPeople[2].ID := 3;
  FPeople[2].Name := 'Carol White';
  FPeople[2].Email := 'carol@example.com';
  FPeople[2].Age := 22;
  FPeople[2].Active := False;

  FPeople[3].ID := 4;
  FPeople[3].Name := 'David Brown';
  FPeople[3].Email := 'david@example.com';
  FPeople[3].Age := 45;
  FPeople[3].Active := True;

  FPeople[4].ID := 5;
  FPeople[4].Name := 'Eva Green';
  FPeople[4].Email := 'eva@example.com';
  FPeople[4].Age := 31;
  FPeople[4].Active := True;

  FPeople[5].ID := 6;
  FPeople[5].Name := 'Frank Black';
  FPeople[5].Email := 'frank@example.com';
  FPeople[5].Age := 52;
  FPeople[5].Active := False;

  FPeople[6].ID := 7;
  FPeople[6].Name := 'Grace Lee';
  FPeople[6].Email := 'grace@example.com';
  FPeople[6].Age := 27;
  FPeople[6].Active := True;

  FPeople[7].ID := 8;
  FPeople[7].Name := 'Henry Ford';
  FPeople[7].Email := 'henry@example.com';
  FPeople[7].Age := 39;
  FPeople[7].Active := True;

  FilterData;
  SortData;
end;

procedure TfrmData.FilterData;
var
  I, Count: Integer;
  SearchText: string;
begin
  SearchText := LowerCase(FEdtSearch.Text);
  Count := 0;
  for I := 0 to High(FPeople) do
  begin
    if FChkShowActiveOnly.Checked and not FPeople[I].Active then
      Continue;
    if (SearchText <> '') and (Pos(SearchText, LowerCase(FPeople[I].Name)) = 0) then
      Continue;
    SetLength(FFiltered, Count + 1);
    FFiltered[Count] := I;
    Count := Count + 1;
  end;
end;

procedure TfrmData.SortData;
var
  I, J, Temp: Integer;
begin
  for I := 0 to High(FFiltered) - 1 do
    for J := I + 1 to High(FFiltered) do
    begin
      case FCmbSortBy.ItemIndex of
        0: if FPeople[FFiltered[I]].ID > FPeople[FFiltered[J]].ID then
           begin
             Temp := FFiltered[I];
             FFiltered[I] := FFiltered[J];
             FFiltered[J] := Temp;
           end;
        1: if FPeople[FFiltered[I]].Name > FPeople[FFiltered[J]].Name then
           begin
             Temp := FFiltered[I];
             FFiltered[I] := FFiltered[J];
             FFiltered[J] := Temp;
           end;
        2: if FPeople[FFiltered[I]].Age > FPeople[FFiltered[J]].Age then
           begin
             Temp := FFiltered[I];
             FFiltered[I] := FFiltered[J];
             FFiltered[J] := Temp;
           end;
      end;
    end;
end;

procedure TfrmData.RefreshGrid;
var
  I: Integer;
begin
  FSgData.RowCount := Length(FFiltered) + 1;
  for I := 0 to High(FFiltered) do
  begin
    FSgData.Cells[0, I + 1] := IntToStr(FPeople[FFiltered[I]].ID);
    FSgData.Cells[1, I + 1] := FPeople[FFiltered[I]].Name;
    FSgData.Cells[2, I + 1] := FPeople[FFiltered[I]].Email;
    FSgData.Cells[3, I + 1] := IntToStr(FPeople[FFiltered[I]].Age);
    if FPeople[FFiltered[I]].Active then
      FSgData.Cells[4, I + 1] := 'Yes'
    else
      FSgData.Cells[4, I + 1] := 'No';
  end;
  FLabelCount.Caption := 'Total records: ' + IntToStr(Length(FPeople)) +
                         ' (showing ' + IntToStr(Length(FFiltered)) + ')';
end;

procedure TfrmData.ShowDetails(Index: Integer);
var
  P: TPerson;
begin
  if (Index < 0) or (Index >= Length(FFiltered)) then Exit;
  P := FPeople[FFiltered[Index]];
  FLstDetails.Clear;
  FLstDetails.Items.Add('ID: ' + IntToStr(P.ID));
  FLstDetails.Items.Add('Name: ' + P.Name);
  FLstDetails.Items.Add('Email: ' + P.Email);
  FLstDetails.Items.Add('Age: ' + IntToStr(P.Age));
  FLstDetails.Items.Add('Status: ' + BoolToStr(P.Active));
  FLabelSelected.Caption := 'Selected: ' + P.Name;
  FCurrentIndex := Index;
  UpdateNavButtons;
end;

procedure TfrmData.UpdateNavButtons;
begin
  FBtnFirst.Enabled := FCurrentIndex > 0;
  FBtnPrev.Enabled := FCurrentIndex > 0;
  FBtnNext.Enabled := FCurrentIndex < Length(FFiltered) - 1;
  FBtnLast.Enabled := FCurrentIndex < Length(FFiltered) - 1;
end;

procedure TfrmData.FormCreate(Sender: TObject);
begin
  BuildControls;
  LoadSampleData;
  RefreshGrid;
  FCurrentIndex := -1;
  FSelectedRecord := '';
end;

procedure TfrmData.FormDestroy(Sender: TObject);
begin
  // Cleanup
end;

procedure TfrmData.sgDataSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
begin
  if ARow > 0 then
    ShowDetails(ARow - 1);
end;

procedure TfrmData.edtSearchChange(Sender: TObject);
begin
  FilterData;
  SortData;
  RefreshGrid;
end;

procedure TfrmData.chkShowActiveOnlyChange(Sender: TObject);
begin
  FilterData;
  SortData;
  RefreshGrid;
end;

procedure TfrmData.cmbSortByChange(Sender: TObject);
begin
  SortData;
  RefreshGrid;
end;

procedure TfrmData.rgViewModeClick(Sender: TObject);
begin
  // View mode changed
end;

procedure TfrmData.lstDetailsClick(Sender: TObject);
begin
  // Detail item clicked
end;

procedure TfrmData.pgDetailsChange(Sender: TObject);
begin
  // Tab changed
end;

procedure TfrmData.btnAddClick(Sender: TObject);
var
  NewPerson: TPerson;
  Len: Integer;
begin
  Len := Length(FPeople);
  SetLength(FPeople, Len + 1);
  NewPerson.ID := GenerateID;
  NewPerson.Name := 'New Person ' + IntToStr(NewPerson.ID);
  NewPerson.Email := 'new' + IntToStr(NewPerson.ID) + '@example.com';
  NewPerson.Age := 25;
  NewPerson.Active := True;
  FPeople[Len] := NewPerson;
  FilterData;
  SortData;
  RefreshGrid;
end;

procedure TfrmData.btnEditClick(Sender: TObject);
begin
  if FCurrentIndex >= 0 then
  begin
    FPeople[FFiltered[FCurrentIndex]].Name := FPeople[FFiltered[FCurrentIndex]].Name + ' (edited)';
    RefreshGrid;
    ShowDetails(FCurrentIndex);
  end;
end;

procedure TfrmData.btnDeleteClick(Sender: TObject);
var
  I, Idx: Integer;
begin
  if FCurrentIndex < 0 then Exit;
  if MessageDlg('Confirm', 'Delete selected record?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;

  Idx := FFiltered[FCurrentIndex];
  for I := Idx to High(FPeople) - 1 do
    FPeople[I] := FPeople[I + 1];
  SetLength(FPeople, Length(FPeople) - 1);

  FilterData;
  SortData;
  RefreshGrid;
  FCurrentIndex := -1;
  FLstDetails.Clear;
  FLabelSelected.Caption := 'Selected: none';
end;

procedure TfrmData.btnSearchClick(Sender: TObject);
begin
  edtSearchChange(Sender);
end;

procedure TfrmData.btnRefreshClick(Sender: TObject);
begin
  FilterData;
  SortData;
  RefreshGrid;
end;

procedure TfrmData.btnExportClick(Sender: TObject);
begin
  ShowMessage('Export: ' + IntToStr(Length(FFiltered)) + ' records would be exported.');
end;

procedure TfrmData.btnSelectClick(Sender: TObject);
begin
  if FCurrentIndex >= 0 then
    FSelectedRecord := FPeople[FFiltered[FCurrentIndex]].Name;
  ModalResult := mrOK;
end;

procedure TfrmData.btnCloseClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TfrmData.btnFirstClick(Sender: TObject);
begin
  if Length(FFiltered) > 0 then
  begin
    FSgData.Row := 1;
    ShowDetails(0);
  end;
end;

procedure TfrmData.btnPrevClick(Sender: TObject);
begin
  if FCurrentIndex > 0 then
  begin
    FSgData.Row := FCurrentIndex;
    ShowDetails(FCurrentIndex - 1);
  end;
end;

procedure TfrmData.btnNextClick(Sender: TObject);
begin
  if FCurrentIndex < Length(FFiltered) - 1 then
  begin
    FSgData.Row := FCurrentIndex + 2;
    ShowDetails(FCurrentIndex + 1);
  end;
end;

procedure TfrmData.btnLastClick(Sender: TObject);
begin
  if Length(FFiltered) > 0 then
  begin
    FSgData.Row := Length(FFiltered);
    ShowDetails(Length(FFiltered) - 1);
  end;
end;

end.
