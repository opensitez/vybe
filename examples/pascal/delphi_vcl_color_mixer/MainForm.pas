unit MainForm;

interface

uses
  SysUtils, Classes, Forms, Controls, StdCtrls, ComCtrls, ExtCtrls, Graphics;

type
  TColorMixerForm = class(TForm)
  private
    FRedTrack: TTrackBar;
    FGreenTrack: TTrackBar;
    FBlueTrack: TTrackBar;
    FPreviewPanel: TPanel;
    FHexLabel: TLabel;
    FRandomButton: TButton;
    procedure BuildUi;
    procedure TrackChange(Sender: TObject);
    procedure RandomClick(Sender: TObject);
    procedure UpdateColor;
  public
    constructor Create(AOwner: TComponent); override;
  end;

var
  ColorMixerForm: TColorMixerForm;

implementation

{$R *.dfm}

constructor TColorMixerForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Randomize;
  BuildUi;
  UpdateColor;
end;

procedure TColorMixerForm.BuildUi;
  procedure SetupTrack(var Track: TTrackBar; const Caption: string; Y: Integer);
  var
    L: TLabel;
  begin
    L := TLabel.Create(Self);
    L.Parent := Self;
    L.Left := 16;
    L.Top := Y;
    L.Caption := Caption;

    Track := TTrackBar.Create(Self);
    Track.Parent := Self;
    Track.Left := 16;
    Track.Top := Y + 16;
    Track.Width := 340;
    Track.Min := 0;
    Track.Max := 255;
    Track.Position := 128;
    Track.OnChange := TrackChange;
  end;
begin
  Caption := 'Color Mixer Studio';
  Width := 390;
  Height := 430;
  Position := poScreenCenter;

  SetupTrack(FRedTrack, 'Red', 16);
  SetupTrack(FGreenTrack, 'Green', 96);
  SetupTrack(FBlueTrack, 'Blue', 176);

  FPreviewPanel := TPanel.Create(Self);
  FPreviewPanel.Parent := Self;
  FPreviewPanel.Left := 16;
  FPreviewPanel.Top := 258;
  FPreviewPanel.Width := 340;
  FPreviewPanel.Height := 90;
  FPreviewPanel.Caption := '';

  FHexLabel := TLabel.Create(Self);
  FHexLabel.Parent := Self;
  FHexLabel.Left := 16;
  FHexLabel.Top := 358;
  FHexLabel.Font.Size := 11;

  FRandomButton := TButton.Create(Self);
  FRandomButton.Parent := Self;
  FRandomButton.Left := 206;
  FRandomButton.Top := 354;
  FRandomButton.Width := 150;
  FRandomButton.Height := 30;
  FRandomButton.Caption := 'Random Palette';
  FRandomButton.OnClick := RandomClick;
end;

procedure TColorMixerForm.TrackChange(Sender: TObject);
begin
  UpdateColor;
end;

procedure TColorMixerForm.RandomClick(Sender: TObject);
begin
  FRedTrack.Position := Random(256);
  FGreenTrack.Position := Random(256);
  FBlueTrack.Position := Random(256);
  UpdateColor;
end;

procedure TColorMixerForm.UpdateColor;
var
  R, G, B: Integer;
  ColorValue: TColor;
begin
  R := FRedTrack.Position;
  G := FGreenTrack.Position;
  B := FBlueTrack.Position;
  ColorValue := RGB(R, G, B);
  FPreviewPanel.Color := ColorValue;
  FHexLabel.Caption := Format('RGB(%d, %d, %d)   Hex: #%2.2x%2.2x%2.2x', [R, G, B, R, G, B]);
end;

end.
