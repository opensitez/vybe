use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 12: Class Properties & Getter/Setter Accessors
// ═══════════════════════════════════════════════════════════

#[test]
fn test_property_direct_field_access() {
    let out = run_pascal(r#"
program Test;
type TItem = class
  private FValue: Integer;
  public property Value: Integer read FValue write FValue;
end;
var item: TItem;
begin
  item := TItem.Create;
  item.Value := 100;
  WriteLn(item.Value);
  item.Free;
end.
"#);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_property_method_accessors() {
    let out = run_pascal(r#"
program Test;
type TAccount = class
  private FBalance: Real;
  private function GetBalance: Real;
  private procedure SetBalance(v: Real);
  public property Balance: Real read GetBalance write SetBalance;
end;
function TAccount.GetBalance: Real; begin Result := FBalance; end;
procedure TAccount.SetBalance(v: Real); begin FBalance := v; end;
var acc: TAccount;
begin
  acc := TAccount.Create;
  acc.Balance := 250.50;
  WriteLn(acc.Balance);
  acc.Free;
end.
"#);
    assert_eq!(out, vec!["250.5"]);
}

#[test]
fn test_property_read_only() {
    let out = run_pascal(r#"
program Test;
type TSystemInfo = class
  private FOSName: String;
  public constructor Create;
  public property OSName: String read FOSName;
end;
constructor TSystemInfo.Create; begin FOSName := 'PascalOS'; end;
var info: TSystemInfo;
begin
  info := TSystemInfo.Create;
  WriteLn(info.OSName);
  info.Free;
end.
"#);
    assert_eq!(out, vec!["PascalOS"]);
}

#[test]
fn test_property_setter_validation_side_effect() {
    let out = run_pascal(r#"
program Test;
type TTemperature = class
  private FCelsius: Integer;
  private procedure SetCelsius(v: Integer);
  public property Celsius: Integer read FCelsius write SetCelsius;
end;
procedure TTemperature.SetCelsius(v: Integer);
begin
  if v < -273 then FCelsius := -273
  else FCelsius := v;
end;
var t: TTemperature;
begin
  t := TTemperature.Create;
  t.Celsius := -300;
  WriteLn(t.Celsius);
  t.Celsius := 25;
  WriteLn(t.Celsius);
  t.Free;
end.
"#);
    assert_eq!(out, vec!["-273", "25"]);
}

#[test]
fn test_property_computed_getter() {
    let out = run_pascal(r#"
program Test;
type TRectangle = class
  private FWidth, FHeight: Integer;
  private function GetArea: Integer;
  public constructor Create(W, H: Integer);
  public property Area: Integer read GetArea;
end;
constructor TRectangle.Create(W, H: Integer); begin FWidth := W; FHeight := H; end;
function TRectangle.GetArea: Integer; begin Result := FWidth * FHeight; end;
var r: TRectangle;
begin
  r := TRectangle.Create(5, 10);
  WriteLn(r.Area);
  r.Free;
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_property_index_specifier() {
    let out = run_pascal(r#"
program Test;
type TDataHolder = class
  private FVal1, FVal2: Integer;
  private function GetVal(index: Integer): Integer;
  private procedure SetVal(index: Integer; value: Integer);
  public property Val1: Integer index 1 read GetVal write SetVal;
  public property Val2: Integer index 2 read GetVal write SetVal;
end;
function TDataHolder.GetVal(index: Integer): Integer;
begin
  if index = 1 then Result := FVal1 else Result := FVal2;
end;
procedure TDataHolder.SetVal(index: Integer; value: Integer);
begin
  if index = 1 then FVal1 := value else FVal2 := value;
end;
var h: TDataHolder;
begin
  h := TDataHolder.Create;
  h.Val1 := 10;
  h.Val2 := 20;
  WriteLn(h.Val1);
  WriteLn(h.Val2);
  h.Free;
end.
"#);
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_property_boolean_flag() {
    let out = run_pascal(r#"
program Test;
type TToggle = class
  private FEnabled: Boolean;
  public property Enabled: Boolean read FEnabled write FEnabled;
end;
var t: TToggle;
begin
  t := TToggle.Create;
  t.Enabled := True;
  WriteLn(t.Enabled);
  t.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_property_enum_type() {
    let out = run_pascal(r#"
program Test;
type TState = (stInactive, stActive);
type TComponent = class
  private FState: TState;
  public property State: TState read FState write FState;
end;
var c: TComponent;
begin
  c := TComponent.Create;
  c.State := stActive;
  WriteLn(Ord(c.State));
  c.Free;
end.
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_property_record_type() {
    let out = run_pascal(r#"
program Test;
type TPoint = record X, Y: Integer; end;
type TShape = class
  private FPos: TPoint;
  public property Pos: TPoint read FPos write FPos;
end;
var s: TShape; p: TPoint;
begin
  s := TShape.Create;
  p.X := 15; p.Y := 30;
  s.Pos := p;
  WriteLn(s.Pos.X);
  WriteLn(s.Pos.Y);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["15", "30"]);
}

#[test]
fn test_property_virtual_getter_override() {
    let out = run_pascal(r#"
program Test;
type TBase = class
  protected function GetTitle: String; virtual;
  public property Title: String read GetTitle;
end;
type TDerived = class(TBase)
  protected function GetTitle: String; override;
end;
function TBase.GetTitle: String; begin Result := 'BaseTitle'; end;
function TDerived.GetTitle: String; begin Result := 'DerivedTitle'; end;
var b: TBase;
begin
  b := TDerived.Create;
  WriteLn(b.Title);
  b.Free;
end.
"#);
    assert_eq!(out, vec!["DerivedTitle"]);
}

#[test]
fn test_property_write_only() {
    let out = run_pascal(r#"
program Test;
type TSecretStore = class
  private FSecret: String;
  private procedure SetSecret(s: String);
  public property Secret: String write SetSecret;
  public function IsSecretSet: Boolean;
end;
procedure TSecretStore.SetSecret(s: String); begin FSecret := s; end;
function TSecretStore.IsSecretSet: Boolean; begin Result := Length(FSecret) > 0; end;
var store: TSecretStore;
begin
  store := TSecretStore.Create;
  store.Secret := 'P@ssword';
  WriteLn(store.IsSecretSet);
  store.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_property_class_property_static() {
    let out = run_pascal(r#"
program Test;
type TGlobalConfig = class
  private class var FAppName: String;
  public class property AppName: String read FAppName write FAppName;
end;
begin
  TGlobalConfig.AppName := 'VybeApp';
  WriteLn(TGlobalConfig.AppName);
end.
"#);
    assert_eq!(out, vec!["VybeApp"]);
}

#[test]
fn test_property_internal_self_access() {
    let out = run_pascal(r#"
program Test;
type TCounter = class
  private FCount: Integer;
  public property Count: Integer read FCount write FCount;
  public procedure Increment;
end;
procedure TCounter.Increment;
begin
  Count := Count + 1;
end;
var c: TCounter;
begin
  c := TCounter.Create;
  c.Count := 10;
  c.Increment;
  WriteLn(c.Count);
  c.Free;
end.
"#);
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_property_subrange_type() {
    let out = run_pascal(r#"
program Test;
type TLevel = 1..5;
type TPlayer = class
  private FLevel: TLevel;
  public property Level: TLevel read FLevel write FLevel;
end;
var p: TPlayer;
begin
  p := TPlayer.Create;
  p.Level := 4;
  WriteLn(p.Level);
  p.Free;
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_property_redeclared_in_derived_class() {
    let out = run_pascal(r#"
program Test;
type TBase = class
  protected FVal: Integer;
  public property Val: Integer read FVal;
end;
type TDerived = class(TBase)
  public property Val: Integer read FVal write FVal;
end;
var d: TDerived;
begin
  d := TDerived.Create;
  d.Val := 999;
  WriteLn(d.Val);
  d.Free;
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_property_string_trim_on_set() {
    let out = run_pascal(r#"
program Test;
type TCleanText = class
  private FText: String;
  private procedure SetText(v: String);
  public property Text: String read FText write SetText;
end;
procedure TCleanText.SetText(v: String);
begin
  FText := Trim(v);
end;
var ct: TCleanText;
begin
  ct := TCleanText.Create;
  ct.Text := '   Cleaned   ';
  WriteLn('[' + ct.Text + ']');
  ct.Free;
end.
"#);
    assert_eq!(out, vec!["[Cleaned]"]);
}

#[test]
fn test_property_chained_modification() {
    let out = run_pascal(r#"
program Test;
type TBox = class
  private FWidth, FHeight: Integer;
  public property Width: Integer read FWidth write FWidth;
  public property Height: Integer read FHeight write FHeight;
  public procedure Scale(factor: Integer);
end;
procedure TBox.Scale(factor: Integer);
begin
  Width := Width * factor;
  Height := Height * factor;
end;
var b: TBox;
begin
  b := TBox.Create;
  b.Width := 10; b.Height := 20;
  b.Scale(2);
  WriteLn(b.Width);
  WriteLn(b.Height);
  b.Free;
end.
"#);
    assert_eq!(out, vec!["20", "40"]);
}

#[test]
fn test_property_pointer_type() {
    let out = run_pascal(r#"
program Test;
type TRefHolder = class
  private FDataPtr: PInteger;
  public property DataPtr: PInteger read FDataPtr write FDataPtr;
end;
var holder: TRefHolder; val: Integer;
begin
  val := 404;
  holder := TRefHolder.Create;
  holder.DataPtr := @val;
  WriteLn(holder.DataPtr^);
  holder.Free;
end.
"#);
    assert_eq!(out, vec!["404"]);
}

#[test]
fn test_property_getter_side_effect_count() {
    let out = run_pascal(r#"
program Test;
type TAccessTracker = class
  private FReads: Integer; FValue: String;
  private function GetValue: String;
  public constructor Create;
  public property Value: String read GetValue;
  public property Reads: Integer read FReads;
end;
constructor TAccessTracker.Create; begin FValue := 'Data'; FReads := 0; end;
function TAccessTracker.GetValue: String;
begin
  Inc(FReads);
  Result := FValue;
end;
var t: TAccessTracker; s: String;
begin
  t := TAccessTracker.Create;
  s := t.Value; s := t.Value; s := t.Value;
  WriteLn(t.Reads);
  t.Free;
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_property_default_specifier_metadata() {
    let out = run_pascal(r#"
program Test;
type TControl = class
  private FVisible: Boolean;
  public constructor Create;
  published property Visible: Boolean read FVisible write FVisible default True;
end;
constructor TControl.Create; begin FVisible := True; end;
var ctrl: TControl;
begin
  ctrl := TControl.Create;
  WriteLn(ctrl.Visible);
  ctrl.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}
