unit SettingsForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, ExtCtrls, ComCtrls, Dialogs;

type
  TfrmSettings = class(TForm)
    procedure FormCreate(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnDefaultsClick(Sender: TObject);
    procedure chkDarkModeChange(Sender: TObject);
    procedure chkAutoSaveChange(Sender: TObject);
    procedure chkNotificationsChange(Sender: TObject);
    procedure seFontSizeChange(Sender: TObject);
    procedure trbOpacityChange(Sender: TObject);
    procedure cmbLanguageChange(Sender: TObject);
    procedure cmbThemeChange(Sender: TObject);
    procedure lstShortcutsClick(Sender: TObject);
    procedure rgStartupClick(Sender: TObject);
    procedure btnPickColorClick(Sender: TObject);
  private
    FLabelTitle: TLabel;
    FLabelInfo: TLabel;
    FLabelOpacity: TLabel;
    FLabelFontSize: TLabel;
    FLabelColorPreview: TLabel;
    FPanelAppearance: TPanel;
    FPanelBehavior: TPanel;
    FPanelShortcuts: TPanel;
    FChkDarkMode: TCheckBox;
    FChkAutoSave: TCheckBox;
    FChkNotifications: TCheckBox;
    FChkMinimizeToTray: TCheckBox;
    FSeFontSize: TSpinEdit;
    FTrbOpacity: TTrackBar;
    FCmbLanguage: TComboBox;
    FCmbTheme: TComboBox;
    FLstShortcuts: TListBox;
    FRgStartup: TRadioGroup;
    FColorDialog: TColorDialog;
    FBtnPickColor: TButton;
    FBtnSave: TButton;
    FBtnCancel: TButton;
    FBtnDefaults: TButton;
    FSelectedColor: string;
    procedure BuildControls;
    procedure ApplyDefaults;
  public
    property SelectedTheme: string read FCmbTheme.Text;
    property AutoSave: Boolean read FChkAutoSave.Checked;
    property SelectedColor: string read FSelectedColor;
  end;

implementation

procedure TfrmSettings.BuildControls;
begin
  Caption := 'Settings';
  Width := 520;
  Height := 480;
  Position := poMainFormCenter;
  BorderStyle := bsDialog;

  FLabelTitle := TLabel.Create(Self);
  FLabelTitle.Parent := Self;
  FLabelTitle.Left := 20;
  FLabelTitle.Top := 10;
  FLabelTitle.Caption := 'Application Settings';
  FLabelTitle.Font.Size := 14;
  FLabelTitle.Font.Style := [fsBold];

  FLabelInfo := TLabel.Create(Self);
  FLabelInfo.Parent := Self;
  FLabelInfo.Left := 20;
  FLabelInfo.Top := 35;
  FLabelInfo.Caption := 'Customize appearance, behavior, and shortcuts.';

  FPanelAppearance := TPanel.Create(Self);
  FPanelAppearance.Parent := Self;
  FPanelAppearance.Left := 20;
  FPanelAppearance.Top := 60;
  FPanelAppearance.Width := 230;
  FPanelAppearance.Height := 200;
  FPanelAppearance.Caption := 'Appearance';

  FChkDarkMode := TCheckBox.Create(Self);
  FChkDarkMode.Parent := FPanelAppearance;
  FChkDarkMode.Left := 10;
  FChkDarkMode.Top := 25;
  FChkDarkMode.Caption := 'Dark Mode';
  FChkDarkMode.OnClick := chkDarkModeChange;

  FCmbTheme := TComboBox.Create(Self);
  FCmbTheme.Parent := FPanelAppearance;
  FCmbTheme.Left := 10;
  FCmbTheme.Top := 55;
  FCmbTheme.Width := 200;
  FCmbTheme.Items.Add('Classic Blue');
  FCmbTheme.Items.Add('Modern Dark');
  FCmbTheme.Items.Add('High Contrast');
  FCmbTheme.Items.Add('Pastel');
  FCmbTheme.ItemIndex := 0;
  FCmbTheme.OnChange := cmbThemeChange;

  with TLabel.Create(Self) do
  begin
    Parent := FPanelAppearance;
    Left := 10;
    Top := 85;
    Caption := 'Accent Color:';
  end;

  FLabelColorPreview := TLabel.Create(Self);
  FLabelColorPreview.Parent := FPanelAppearance;
  FLabelColorPreview.Left := 100;
  FLabelColorPreview.Top := 85;
  FLabelColorPreview.Caption := '[Blue]';
  FLabelColorPreview.Font.Color := clBlue;
  FLabelColorPreview.Font.Style := [fsBold];

  FBtnPickColor := TButton.Create(Self);
  FBtnPickColor.Parent := FPanelAppearance;
  FBtnPickColor.Left := 10;
  FBtnPickColor.Top := 110;
  FBtnPickColor.Width := 120;
  FBtnPickColor.Caption := 'Pick Color...';
  FBtnPickColor.OnClick := btnPickColorClick;

  FLabelFontSize := TLabel.Create(Self);
  FLabelFontSize.Parent := FPanelAppearance;
  FLabelFontSize.Left := 10;
  FLabelFontSize.Top := 145;
  FLabelFontSize.Caption := 'Font Size: 10';

  FSeFontSize := TSpinEdit.Create(Self);
  FSeFontSize.Parent := FPanelAppearance;
  FSeFontSize.Left := 10;
  FSeFontSize.Top := 165;
  FSeFontSize.Width := 80;
  FSeFontSize.MinValue := 8;
  FSeFontSize.MaxValue := 24;
  FSeFontSize.Value := 10;
  FSeFontSize.OnChange := seFontSizeChange;

  FPanelBehavior := TPanel.Create(Self);
  FPanelBehavior.Parent := Self;
  FPanelBehavior.Left := 270;
  FPanelBehavior.Top := 60;
  FPanelBehavior.Width := 220;
  FPanelBehavior.Height := 200;
  FPanelBehavior.Caption := 'Behavior';

  FChkAutoSave := TCheckBox.Create(Self);
  FChkAutoSave.Parent := FPanelBehavior;
  FChkAutoSave.Left := 10;
  FChkAutoSave.Top := 25;
  FChkAutoSave.Caption := 'Auto-save on exit';
  FChkAutoSave.OnClick := chkAutoSaveChange;

  FChkNotifications := TCheckBox.Create(Self);
  FChkNotifications.Parent := FPanelBehavior;
  FChkNotifications.Left := 10;
  FChkNotifications.Top := 50;
  FChkNotifications.Caption := 'Show notifications';
  FChkNotifications.Checked := True;
  FChkNotifications.OnClick := chkNotificationsChange;

  FChkMinimizeToTray := TCheckBox.Create(Self);
  FChkMinimizeToTray.Parent := FPanelBehavior;
  FChkMinimizeToTray.Left := 10;
  FChkMinimizeToTray.Top := 75;
  FChkMinimizeToTray.Caption := 'Minimize to tray';

  with TLabel.Create(Self) do
  begin
    Parent := FPanelBehavior;
    Left := 10;
    Top := 105;
    Caption := 'Language:';
  end;

  FCmbLanguage := TComboBox.Create(Self);
  FCmbLanguage.Parent := FPanelBehavior;
  FCmbLanguage.Left := 10;
  FCmbLanguage.Top := 125;
  FCmbLanguage.Width := 180;
  FCmbLanguage.Items.Add('English');
  FCmbLanguage.Items.Add('French');
  FCmbLanguage.Items.Add('German');
  FCmbLanguage.Items.Add('Spanish');
  FCmbLanguage.ItemIndex := 0;
  FCmbLanguage.OnChange := cmbLanguageChange;

  FLabelOpacity := TLabel.Create(Self);
  FLabelOpacity.Parent := FPanelBehavior;
  FLabelOpacity.Left := 10;
  FLabelOpacity.Top := 155;
  FLabelOpacity.Caption := 'Opacity: 100%';

  FTrbOpacity := TTrackBar.Create(Self);
  FTrbOpacity.Parent := FPanelBehavior;
  FTrbOpacity.Left := 10;
  FTrbOpacity.Top := 175;
  FTrbOpacity.Width := 180;
  FTrbOpacity.Min := 50;
  FTrbOpacity.Max := 100;
  FTrbOpacity.Position := 100;
  FTrbOpacity.OnChange := trbOpacityChange;

  FPanelShortcuts := TPanel.Create(Self);
  FPanelShortcuts.Parent := Self;
  FPanelShortcuts.Left := 20;
  FPanelShortcuts.Top := 270;
  FPanelShortcuts.Width := 470;
  FPanelShortcuts.Height := 100;
  FPanelShortcuts.Caption := 'Startup Options';

  FRgStartup := TRadioGroup.Create(Self);
  FRgStartup.Parent := FPanelShortcuts;
  FRgStartup.Left := 10;
  FRgStartup.Top := 20;
  FRgStartup.Width := 440;
  FRgStartup.Height := 70;
  FRgStartup.Caption := '';
  FRgStartup.Items.Add('Open blank workspace');
  FRgStartup.Items.Add('Restore last session');
  FRgStartup.Items.Add('Show welcome dialog');
  FRgStartup.ItemIndex := 0;
  FRgStartup.OnClick := rgStartupClick;

  FBtnSave := TButton.Create(Self);
  FBtnSave.Parent := Self;
  FBtnSave.Left := 220;
  FBtnSave.Top := 390;
  FBtnSave.Width := 80;
  FBtnSave.Height := 30;
  FBtnSave.Caption := 'Save';
  FBtnSave.ModalResult := mrOK;
  FBtnSave.OnClick := btnSaveClick;

  FBtnCancel := TButton.Create(Self);
  FBtnCancel.Parent := Self;
  FBtnCancel.Left := 310;
  FBtnCancel.Top := 390;
  FBtnCancel.Width := 80;
  FBtnCancel.Height := 30;
  FBtnCancel.Caption := 'Cancel';
  FBtnCancel.ModalResult := mrCancel;
  FBtnCancel.OnClick := btnCancelClick;

  FBtnDefaults := TButton.Create(Self);
  FBtnDefaults.Parent := Self;
  FBtnDefaults.Left := 400;
  FBtnDefaults.Top := 390;
  FBtnDefaults.Width := 90;
  FBtnDefaults.Height := 30;
  FBtnDefaults.Caption := 'Defaults';
  FBtnDefaults.OnClick := btnDefaultsClick;

  FColorDialog := TColorDialog.Create(Self);
  FSelectedColor := 'Blue';
end;

procedure TfrmSettings.FormCreate(Sender: TObject);
begin
  BuildControls;
end;

procedure TfrmSettings.ApplyDefaults;
begin
  FChkDarkMode.Checked := False;
  FChkAutoSave.Checked := False;
  FChkNotifications.Checked := True;
  FChkMinimizeToTray.Checked := False;
  FCmbTheme.ItemIndex := 0;
  FCmbLanguage.ItemIndex := 0;
  FSeFontSize.Value := 10;
  FTrbOpacity.Position := 100;
  FRgStartup.ItemIndex := 0;
  FLabelOpacity.Caption := 'Opacity: 100%';
  FLabelFontSize.Caption := 'Font Size: 10';
  FLabelColorPreview.Caption := '[Blue]';
  FLabelColorPreview.Font.Color := clBlue;
  FSelectedColor := 'Blue';
end;

procedure TfrmSettings.btnSaveClick(Sender: TObject);
begin
  if FChkNotifications.Checked then
    ShowMessage('Settings saved successfully!');
  ModalResult := mrOK;
end;

procedure TfrmSettings.btnCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TfrmSettings.btnDefaultsClick(Sender: TObject);
begin
  if MessageDlg('Reset', 'Restore all default settings?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    ApplyDefaults;
end;

procedure TfrmSettings.chkDarkModeChange(Sender: TObject);
begin
  if FChkDarkMode.Checked then
    Color := clBlack
  else
    Color := clBtnFace;
end;

procedure TfrmSettings.chkAutoSaveChange(Sender: TObject);
begin
  // Auto-save setting changed
end;

procedure TfrmSettings.chkNotificationsChange(Sender: TObject);
begin
  // Notifications setting changed
end;

procedure TfrmSettings.seFontSizeChange(Sender: TObject);
begin
  FLabelFontSize.Caption := 'Font Size: ' + IntToStr(FSeFontSize.Value);
end;

procedure TfrmSettings.trbOpacityChange(Sender: TObject);
begin
  FLabelOpacity.Caption := 'Opacity: ' + IntToStr(FTrbOpacity.Position) + '%';
end;

procedure TfrmSettings.cmbLanguageChange(Sender: TObject);
begin
  // Language changed
end;

procedure TfrmSettings.cmbThemeChange(Sender: TObject);
begin
  // Theme changed
end;

procedure TfrmSettings.lstShortcutsClick(Sender: TObject);
begin
  // Shortcut selected
end;

procedure TfrmSettings.rgStartupClick(Sender: TObject);
begin
  // Startup option changed
end;

procedure TfrmSettings.btnPickColorClick(Sender: TObject);
begin
  if FColorDialog.Execute then
  begin
    FSelectedColor := 'Custom';
    FLabelColorPreview.Caption := '[Custom]';
    FLabelColorPreview.Font.Color := FColorDialog.Color;
  end;
end;

end.
