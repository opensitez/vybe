unit MainForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, ExtCtrls, Menus, Dialogs;

type
  TfrmMain = class(TForm)
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOpenSettingsClick(Sender: TObject);
    procedure btnOpenDataClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnCalculateClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure btnAddItemClick(Sender: TObject);
    procedure btnRemoveItemClick(Sender: TObject);
    procedure lstItemsClick(Sender: TObject);
    procedure chkEnableExtraChange(Sender: TObject);
    procedure rgOperationClick(Sender: TObject);
    procedure tbSliderChange(Sender: TObject);
    procedure cmbCategoryChange(Sender: TObject);
    procedure edtInputKeyPress(Sender: TObject; var Key: Char);
    procedure TimerTick(Sender: TObject);
    procedure mnuFileExitClick(Sender: TObject);
    procedure mnuHelpAboutClick(Sender: TObject);
    procedure mnuToolsResetClick(Sender: TObject);
  private
    FTimer: TTimer;
    FStatusPanel: TPanel;
    FMainMenu: TMainMenu;
    FMenuFile, FMenuTools, FMenuHelp: TMenuItem;
    FLabelTitle: TLabel;
    FLabelStatus: TLabel;
    FLabelResult: TLabel;
    FLabelSlider: TLabel;
    FEdtInput: TEdit;
    FEdtMultiplier: TEdit;
    FBtnCalculate: TButton;
    FBtnClear: TButton;
    FBtnAddItem: TButton;
    FBtnRemoveItem: TButton;
    FBtnOpenSettings: TButton;
    FBtnOpenData: TButton;
    FBtnExit: TButton;
    FLstItems: TListBox;
    FChkEnableExtra: TCheckBox;
    FRgOperation: TRadioGroup;
    FTbSlider: TTrackBar;
    FCmbCategory: TComboBox;
    FMemoLog: TMemo;
    FPanelButtons: TPanel;
    FPanelInput: TPanel;
    FPanelResult: TPanel;
    FItemCount: Integer;
    procedure BuildMenu;
    procedure BuildControls;
    procedure UpdateStatus(const Msg: string);
    function DoCalculation(A, B: Real; Op: Integer): Real;
  public
  end;

var
  frmMain: TfrmMain;

implementation

uses
  SettingsForm, DataForm;

procedure TfrmMain.BuildMenu;
begin
  FMainMenu := TMainMenu.Create(Self);

  FMenuFile := TMenuItem.Create(FMainMenu);
  FMenuFile.Caption := '&File';
  FMainMenu.Items.Add(FMenuFile);

  with TMenuItem.Create(FMainMenu) do
  begin
    Caption := 'E&xit';
    OnClick := mnuFileExitClick;
    FMenuFile.Add(Self);
  end;

  FMenuTools := TMenuItem.Create(FMainMenu);
  FMenuTools.Caption := '&Tools';
  FMainMenu.Items.Add(FMenuTools);

  with TMenuItem.Create(FMainMenu) do
  begin
    Caption := '&Reset All';
    OnClick := mnuToolsResetClick;
    FMenuTools.Add(Self);
  end;

  FMenuHelp := TMenuItem.Create(FMainMenu);
  FMenuHelp.Caption := '&Help';
  FMainMenu.Items.Add(FMenuHelp);

  with TMenuItem.Create(FMainMenu) do
  begin
    Caption := '&About';
    OnClick := mnuHelpAboutClick;
    FMenuHelp.Add(Self);
  end;
end;

procedure TfrmMain.BuildControls;
begin
  Caption := 'Delphi Multi-Form Demo - Main';
  Width := 700;
  Height := 550;
  Position := poScreenCenter;

  FLabelTitle := TLabel.Create(Self);
  FLabelTitle.Parent := Self;
  FLabelTitle.Left := 20;
  FLabelTitle.Top := 10;
  FLabelTitle.Caption := 'Calculator & List Manager';
  FLabelTitle.Font.Size := 14;
  FLabelTitle.Font.Style := [fsBold];

  FPanelInput := TPanel.Create(Self);
  FPanelInput.Parent := Self;
  FPanelInput.Left := 20;
  FPanelInput.Top := 40;
  FPanelInput.Width := 320;
  FPanelInput.Height := 200;
  FPanelInput.Caption := '';

  with TLabel.Create(Self) do
  begin
    Parent := FPanelInput;
    Left := 10;
    Top := 10;
    Caption := 'Input Value:';
  end;

  FEdtInput := TEdit.Create(Self);
  FEdtInput.Parent := FPanelInput;
  FEdtInput.Left := 10;
  FEdtInput.Top := 30;
  FEdtInput.Width := 120;
  FEdtInput.Text := '10';
  FEdtInput.OnKeyPress := edtInputKeyPress;

  with TLabel.Create(Self) do
  begin
    Parent := FPanelInput;
    Left := 150;
    Top := 10;
    Caption := 'Multiplier:';
  end;

  FEdtMultiplier := TEdit.Create(Self);
  FEdtMultiplier.Parent := FPanelInput;
  FEdtMultiplier.Left := 150;
  FEdtMultiplier.Top := 30;
  FEdtMultiplier.Width := 80;
  FEdtMultiplier.Text := '2.5';

  FRgOperation := TRadioGroup.Create(Self);
  FRgOperation.Parent := FPanelInput;
  FRgOperation.Left := 10;
  FRgOperation.Top := 65;
  FRgOperation.Width := 140;
  FRgOperation.Height := 90;
  FRgOperation.Caption := 'Operation';
  FRgOperation.Items.Add('Add');
  FRgOperation.Items.Add('Subtract');
  FRgOperation.Items.Add('Multiply');
  FRgOperation.Items.Add('Divide');
  FRgOperation.ItemIndex := 2;
  FRgOperation.OnClick := rgOperationClick;

  FChkEnableExtra := TCheckBox.Create(Self);
  FChkEnableExtra.Parent := FPanelInput;
  FChkEnableExtra.Left := 170;
  FChkEnableExtra.Top := 70;
  FChkEnableExtra.Caption := 'Enable Bonus';
  FChkEnableExtra.OnClick := chkEnableExtraChange;

  FCmbCategory := TComboBox.Create(Self);
  FCmbCategory.Parent := FPanelInput;
  FCmbCategory.Left := 170;
  FCmbCategory.Top := 100;
  FCmbCategory.Width := 120;
  FCmbCategory.Items.Add('Standard');
  FCmbCategory.Items.Add('Scientific');
  FCmbCategory.Items.Add('Financial');
  FCmbCategory.ItemIndex := 0;
  FCmbCategory.OnChange := cmbCategoryChange;

  FPanelResult := TPanel.Create(Self);
  FPanelResult.Parent := Self;
  FPanelResult.Left := 20;
  FPanelResult.Top := 250;
  FPanelResult.Width := 320;
  FPanelResult.Height := 80;
  FPanelResult.Caption := '';

  FLabelResult := TLabel.Create(Self);
  FLabelResult.Parent := FPanelResult;
  FLabelResult.Left := 10;
  FLabelResult.Top := 10;
  FLabelResult.Caption := 'Result: (none)';
  FLabelResult.Font.Size := 12;
  FLabelResult.Font.Color := clBlue;

  FLabelSlider := TLabel.Create(Self);
  FLabelSlider.Parent := FPanelResult;
  FLabelSlider.Left := 10;
  FLabelSlider.Top := 40;
  FLabelSlider.Caption := 'Precision: 50%';

  FTbSlider := TTrackBar.Create(Self);
  FTbSlider.Parent := FPanelResult;
  FTbSlider.Left := 120;
  FTbSlider.Top := 40;
  FTbSlider.Width := 180;
  FTbSlider.Min := 0;
  FTbSlider.Max := 100;
  FTbSlider.Position := 50;
  FTbSlider.OnChange := tbSliderChange;

  FPanelButtons := TPanel.Create(Self);
  FPanelButtons.Parent := Self;
  FPanelButtons.Left := 360;
  FPanelButtons.Top := 40;
  FPanelButtons.Width := 300;
  FPanelButtons.Height := 290;
  FPanelButtons.Caption := '';

  FBtnCalculate := TButton.Create(Self);
  FBtnCalculate.Parent := FPanelButtons;
  FBtnCalculate.Left := 10;
  FBtnCalculate.Top := 10;
  FBtnCalculate.Width := 120;
  FBtnCalculate.Height := 35;
  FBtnCalculate.Caption := 'Calculate';
  FBtnCalculate.OnClick := btnCalculateClick;

  FBtnClear := TButton.Create(Self);
  FBtnClear.Parent := FPanelButtons;
  FBtnClear.Left := 140;
  FBtnClear.Top := 10;
  FBtnClear.Width := 120;
  FBtnClear.Height := 35;
  FBtnClear.Caption := 'Clear';
  FBtnClear.OnClick := btnClearClick;

  FBtnAddItem := TButton.Create(Self);
  FBtnAddItem.Parent := FPanelButtons;
  FBtnAddItem.Left := 10;
  FBtnAddItem.Top := 55;
  FBtnAddItem.Width := 120;
  FBtnAddItem.Height := 35;
  FBtnAddItem.Caption := 'Add Item';
  FBtnAddItem.OnClick := btnAddItemClick;

  FBtnRemoveItem := TButton.Create(Self);
  FBtnRemoveItem.Parent := FPanelButtons;
  FBtnRemoveItem.Left := 140;
  FBtnRemoveItem.Top := 55;
  FBtnRemoveItem.Width := 120;
  FBtnRemoveItem.Height := 35;
  FBtnRemoveItem.Caption := 'Remove Item';
  FBtnRemoveItem.OnClick := btnRemoveItemClick;

  FBtnOpenSettings := TButton.Create(Self);
  FBtnOpenSettings.Parent := FPanelButtons;
  FBtnOpenSettings.Left := 10;
  FBtnOpenSettings.Top := 110;
  FBtnOpenSettings.Width := 250;
  FBtnOpenSettings.Height := 40;
  FBtnOpenSettings.Caption := 'Open Settings Form';
  FBtnOpenSettings.OnClick := btnOpenSettingsClick;

  FBtnOpenData := TButton.Create(Self);
  FBtnOpenData.Parent := FPanelButtons;
  FBtnOpenData.Left := 10;
  FBtnOpenData.Top := 160;
  FBtnOpenData.Width := 250;
  FBtnOpenData.Height := 40;
  FBtnOpenData.Caption := 'Open Data Browser Form';
  FBtnOpenData.OnClick := btnOpenDataClick;

  FBtnExit := TButton.Create(Self);
  FBtnExit.Parent := FPanelButtons;
  FBtnExit.Left := 10;
  FBtnExit.Top := 220;
  FBtnExit.Width := 250;
  FBtnExit.Height := 40;
  FBtnExit.Caption := 'Exit Application';
  FBtnExit.OnClick := btnExitClick;

  FLstItems := TListBox.Create(Self);
  FLstItems.Parent := Self;
  FLstItems.Left := 20;
  FLstItems.Top := 340;
  FLstItems.Width := 320;
  FLstItems.Height := 120;
  FLstItems.OnClick := lstItemsClick;

  FMemoLog := TMemo.Create(Self);
  FMemoLog.Parent := Self;
  FMemoLog.Left := 360;
  FMemoLog.Top := 340;
  FMemoLog.Width := 300;
  FMemoLog.Height := 120;
  FMemoLog.ReadOnly := True;
  FMemoLog.Lines.Add('Application started.');
  FMemoLog.Lines.Add('Ready.');

  FStatusPanel := TPanel.Create(Self);
  FStatusPanel.Parent := Self;
  FStatusPanel.Left := 0;
  FStatusPanel.Top := Height - 60;
  FStatusPanel.Width := Width;
  FStatusPanel.Height := 30;
  FStatusPanel.Align := alBottom;
  FStatusPanel.Caption := '';
  FStatusPanel.BevelOuter := bvLowered;

  FLabelStatus := TLabel.Create(Self);
  FLabelStatus.Parent := FStatusPanel;
  FLabelStatus.Left := 10;
  FLabelStatus.Top := 8;
  FLabelStatus.Caption := 'Status: Ready';

  FTimer := TTimer.Create(Self);
  FTimer.Interval := 1000;
  FTimer.OnTimer := TimerTick;
  FTimer.Enabled := True;

  FItemCount := 0;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  BuildMenu;
  BuildControls;
  UpdateStatus('Main form loaded');
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  FTimer.Free;
end;

procedure TfrmMain.UpdateStatus(const Msg: string);
begin
  FLabelStatus.Caption := 'Status: ' + Msg;
  FMemoLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) + ' - ' + Msg);
end;

function TfrmMain.DoCalculation(A, B: Real; Op: Integer): Real;
begin
  case Op of
    0: Result := A + B;
    1: Result := A - B;
    2: Result := A * B;
    3: begin
         if B <> 0 then
           Result := A / B
         else
         begin
           ShowMessage('Cannot divide by zero!');
           Result := 0;
         end;
       end;
  else
    Result := 0;
  end;

  if FChkEnableExtra.Checked then
    Result := Result + 10;
end;

procedure TfrmMain.btnCalculateClick(Sender: TObject);
var
  A, B, Res: Real;
begin
  A := StrToFloat(FEdtInput.Text);
  B := StrToFloat(FEdtMultiplier.Text);
  Res := DoCalculation(A, B, FRgOperation.ItemIndex);
  FLabelResult.Caption := 'Result: ' + FloatToStrF(Res, ffFixed, 10, FTbSlider.Position div 10);
  UpdateStatus('Calculation performed: ' + FloatToStr(Res));
end;

procedure TfrmMain.btnClearClick(Sender: TObject);
begin
  FEdtInput.Text := '0';
  FEdtMultiplier.Text := '1';
  FLabelResult.Caption := 'Result: (none)';
  FRgOperation.ItemIndex := 0;
  FChkEnableExtra.Checked := False;
  FTbSlider.Position := 50;
  FLabelSlider.Caption := 'Precision: 50%';
  UpdateStatus('Fields cleared');
end;

procedure TfrmMain.btnAddItemClick(Sender: TObject);
begin
  FItemCount := FItemCount + 1;
  FLstItems.Items.Add('Item #' + IntToStr(FItemCount) + ' - ' + FCmbCategory.Text);
  UpdateStatus('Added item #' + IntToStr(FItemCount));
end;

procedure TfrmMain.btnRemoveItemClick(Sender: TObject);
begin
  if FLstItems.ItemIndex >= 0 then
  begin
    FLstItems.Items.Delete(FLstItems.ItemIndex);
    UpdateStatus('Item removed');
  end
  else
    ShowMessage('Please select an item to remove');
end;

procedure TfrmMain.lstItemsClick(Sender: TObject);
begin
  if FLstItems.ItemIndex >= 0 then
    UpdateStatus('Selected: ' + FLstItems.Items[FLstItems.ItemIndex]);
end;

procedure TfrmMain.chkEnableExtraChange(Sender: TObject);
begin
  if FChkEnableExtra.Checked then
    UpdateStatus('Bonus mode enabled')
  else
    UpdateStatus('Bonus mode disabled');
end;

procedure TfrmMain.rgOperationClick(Sender: TObject);
begin
  UpdateStatus('Operation changed to: ' + FRgOperation.Items[FRgOperation.ItemIndex]);
end;

procedure TfrmMain.tbSliderChange(Sender: TObject);
begin
  FLabelSlider.Caption := 'Precision: ' + IntToStr(FTbSlider.Position) + '%';
end;

procedure TfrmMain.cmbCategoryChange(Sender: TObject);
begin
  UpdateStatus('Category changed to: ' + FCmbCategory.Text);
end;

procedure TfrmMain.edtInputKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    btnCalculateClick(Sender);
end;

procedure TfrmMain.TimerTick(Sender: TObject);
begin
  Caption := 'Delphi Multi-Form Demo - Main [' + FormatDateTime('hh:nn:ss', Now) + ']';
end;

procedure TfrmMain.mnuFileExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.mnuHelpAboutClick(Sender: TObject);
begin
  ShowMessage('Delphi Multi-Form Demo' + #13#10 +
              'Version 1.0' + #13#10 +
              'Built with Pascal/Delphi');
end;

procedure TfrmMain.mnuToolsResetClick(Sender: TObject);
begin
  btnClearClick(Sender);
  FLstItems.Clear;
  FItemCount := 0;
  UpdateStatus('All data reset');
end;

procedure TfrmMain.btnOpenSettingsClick(Sender: TObject);
var
  Form: TfrmSettings;
begin
  Form := TfrmSettings.Create(Self);
  try
    Form.ShowModal;
    if Form.ModalResult = mrOK then
      UpdateStatus('Settings saved: Theme=' + Form.SelectedTheme + ', AutoSave=' + BoolToStr(Form.AutoSave));
  finally
    Form.Free;
  end;
end;

procedure TfrmMain.btnOpenDataClick(Sender: TObject);
var
  Form: TfrmData;
begin
  Form := TfrmData.Create(Self);
  try
    Form.ShowModal;
    if Form.SelectedRecord <> '' then
      UpdateStatus('Data selected: ' + Form.SelectedRecord);
  finally
    Form.Free;
  end;
end;

procedure TfrmMain.btnExitClick(Sender: TObject);
begin
  if MessageDlg('Confirm Exit', 'Are you sure you want to exit?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    Close;
end;

end.
