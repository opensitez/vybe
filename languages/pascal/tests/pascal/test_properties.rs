use super::helpers::run_pascal;

#[test]
fn test_property_read_write_basic() {
    let src = r#"
program T;
type
  TBox = class
  private
    FWidth: Integer;
  public
    property Width: Integer read FWidth write FWidth;
  end;
var
  b: TBox;
begin
  b := TBox.Create;
  b.Width := 42;
  WriteLn(b.Width);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_property_with_getter() {
    let src = r#"
program T;
type
  TCircle = class
  private
    FRadius: Integer;
    function GetArea: Integer;
  public
    property Radius: Integer read FRadius write FRadius;
    property Area: Integer read GetArea;
  end;

function TCircle.GetArea: Integer;
begin
  Result := FRadius * FRadius * 3;
end;

var
  c: TCircle;
begin
  c := TCircle.Create;
  c.Radius := 5;
  WriteLn(c.Radius);
  WriteLn(c.Area);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["5", "75"]);
}

#[test]
fn test_property_with_setter() {
    let src = r#"
program T;
type
  TPositive = class
  private
    FVal: Integer;
    procedure SetVal(v: Integer);
  public
    property Val: Integer read FVal write SetVal;
  end;

procedure TPositive.SetVal(v: Integer);
begin
  if v >= 0 then FVal := v
  else FVal := 0;
end;

var
  p: TPositive;
begin
  p := TPositive.Create;
  p.Val := 10;
  WriteLn(p.Val);
  p.Val := -5;
  WriteLn(p.Val);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["10", "0"]);
}

#[test]
fn test_property_chain_dependent() {
    let src = r#"
program T;
type
  TRect = class
  private
    FW, FH: Integer;
    function GetArea: Integer;
    function GetPerimeter: Integer;
  public
    property W: Integer read FW write FW;
    property H: Integer read FH write FH;
    property Area: Integer read GetArea;
    property Perimeter: Integer read GetPerimeter;
  end;

function TRect.GetArea: Integer;
begin
  Result := FW * FH;
end;

function TRect.GetPerimeter: Integer;
begin
  Result := 2 * (FW + FH);
end;

var
  r: TRect;
begin
  r := TRect.Create;
  r.W := 4;
  r.H := 6;
  WriteLn(r.Area);
  WriteLn(r.Perimeter);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["24", "20"]);
}

#[test]
fn test_property_inherited() {
    let src = r#"
program T;
type
  TBase = class
  private
    FID: Integer;
  public
    property ID: Integer read FID write FID;
  end;
  TChild = class(TBase)
  private
    FName: string;
  public
    property Name: string read FName write FName;
  end;
var
  c: TChild;
begin
  c := TChild.Create;
  c.ID := 1;
  c.Name := 'test';
  WriteLn(c.ID);
  WriteLn(c.Name);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1", "test"]);
}

#[test]
fn test_property_setter_validates() {
    let src = r#"
program T;
type
  TRange = class
  private
    FValue: Integer;
    FMin, FMax: Integer;
    procedure SetValue(v: Integer);
  public
    constructor Create(mn, mx: Integer);
    property Value: Integer read FValue write SetValue;
  end;

constructor TRange.Create(mn, mx: Integer);
begin
  inherited Create;
  FMin := mn;
  FMax := mx;
  FValue := mn;
end;

procedure TRange.SetValue(v: Integer);
begin
  if v < FMin then FValue := FMin
  else if v > FMax then FValue := FMax
  else FValue := v;
end;

var
  r: TRange;
begin
  r := TRange.Create(0, 100);
  r.Value := 50;
  WriteLn(r.Value);
  r.Value := -10;
  WriteLn(r.Value);
  r.Value := 200;
  WriteLn(r.Value);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["50", "0", "100"]);
}

#[test]
fn test_property_array_index() {
    let src = r#"
program T;
type
  TVector = class
  private
    FData: array[0..4] of Integer;
    function GetItem(idx: Integer): Integer;
    procedure SetItem(idx, val: Integer);
  public
    property Items[idx: Integer]: Integer read GetItem write SetItem;
  end;

function TVector.GetItem(idx: Integer): Integer;
begin
  Result := FData[idx];
end;

procedure TVector.SetItem(idx, val: Integer);
begin
  FData[idx] := val;
end;

var
  v: TVector;
  i: Integer;
begin
  v := TVector.Create;
  for i := 0 to 4 do
    v.Items[i] := i * 10;
  for i := 0 to 4 do
    WriteLn(v.Items[i]);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["0", "10", "20", "30", "40"]);
}

#[test]
fn test_property_default_index() {
    let src = r#"
program T;
type
  TList = class
  private
    FItems: array[0..2] of string;
    FCount: Integer;
    function GetItem(i: Integer): string;
    procedure SetItem(i: Integer; const v: string);
  public
    property Items[i: Integer]: string read GetItem write SetItem; default;
    property Count: Integer read FCount;
    procedure Add(s: string);
  end;

function TList.GetItem(i: Integer): string;
begin
  Result := FItems[i];
end;

procedure TList.SetItem(i: Integer; const v: string);
begin
  FItems[i] := v;
end;

procedure TList.Add(s: string);
begin
  FItems[FCount] := s;
  FCount := FCount + 1;
end;

var
  lst: TList;
begin
  lst := TList.Create;
  lst.Add('alpha');
  lst.Add('beta');
  lst.Add('gamma');
  WriteLn(lst[0]);
  WriteLn(lst[1]);
  WriteLn(lst[2]);
  WriteLn(lst.Count);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["alpha", "beta", "gamma", "3"]);
}

#[test]
fn test_property_override_in_child() {
    let src = r#"
program T;
type
  TBase = class
  private
    FVal: Integer;
  public
    property Val: Integer read FVal write FVal;
    function Describe: string; virtual;
  end;
  TChild = class(TBase)
    function Describe: string; override;
  end;

function TBase.Describe: string;
begin
  Result := 'base:' + IntToStr(Val);
end;

function TChild.Describe: string;
begin
  Result := 'child:' + IntToStr(Val);
end;

var
  b: TBase;
  c: TChild;
begin
  b := TBase.Create;
  c := TChild.Create;
  b.Val := 10;
  c.Val := 20;
  WriteLn(b.Describe);
  WriteLn(c.Describe);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["base:10", "child:20"]);
}

#[test]
fn test_property_computed_bool() {
    let src = r#"
program T;
type
  TAccount = class
  private
    FBalance: Integer;
    function GetIsPositive: Boolean;
  public
    property Balance: Integer read FBalance write FBalance;
    property IsPositive: Boolean read GetIsPositive;
  end;

function TAccount.GetIsPositive: Boolean;
begin
  Result := FBalance > 0;
end;

var
  acc: TAccount;
begin
  acc := TAccount.Create;
  acc.Balance := 100;
  WriteLn(acc.IsPositive);
  acc.Balance := -50;
  WriteLn(acc.IsPositive);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn test_property_readonly() {
    let src = r#"
program T;
type
  TDate = class
  private
    FDay, FMonth, FYear: Integer;
  public
    constructor Create(d, m, y: Integer);
    property Day: Integer read FDay;
    property Month: Integer read FMonth;
    property Year: Integer read FYear;
    function ToString: string;
  end;

constructor TDate.Create(d, m, y: Integer);
begin
  inherited Create;
  FDay := d;
  FMonth := m;
  FYear := y;
end;

function TDate.ToString: string;
begin
  Result := Format('%02d/%02d/%04d', [FDay, FMonth, FYear]);
end;

var
  dt: TDate;
begin
  dt := TDate.Create(15, 6, 2024);
  WriteLn(dt.Day);
  WriteLn(dt.Month);
  WriteLn(dt.Year);
  WriteLn(dt.ToString);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["15", "6", "2024", "15/06/2024"]);
}

#[test]
fn test_property_in_loop() {
    let src = r#"
program T;
type
  TCounter = class
  private
    FValue: Integer;
    procedure SetValue(v: Integer);
    function GetValue: Integer;
  public
    property Value: Integer read GetValue write SetValue;
    procedure Tick;
  end;

procedure TCounter.SetValue(v: Integer);
begin
  if v >= 0 then FValue := v;
end;

function TCounter.GetValue: Integer;
begin
  Result := FValue;
end;

procedure TCounter.Tick;
begin
  FValue := FValue + 1;
end;

var
  c: TCounter;
  i: Integer;
begin
  c := TCounter.Create;
  c.Value := 0;
  for i := 1 to 5 do
    c.Tick;
  WriteLn(c.Value);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_property_string_manipulation() {
    let src = r#"
program T;
type
  TLabel = class
  private
    FText: string;
    function GetUpper: string;
    function GetLen: Integer;
  public
    property Text: string read FText write FText;
    property Upper: string read GetUpper;
    property Len: Integer read GetLen;
  end;

function TLabel.GetUpper: string;
begin
  Result := UpperCase(FText);
end;

function TLabel.GetLen: Integer;
begin
  Result := Length(FText);
end;

var
  lbl: TLabel;
begin
  lbl := TLabel.Create;
  lbl.Text := 'hello';
  WriteLn(lbl.Text);
  WriteLn(lbl.Upper);
  WriteLn(lbl.Len);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["hello", "HELLO", "5"]);
}

#[test]
fn test_property_with_notification() {
    let src = r#"
program T;
type
  TObservable = class
  private
    FValue: Integer;
    FChanges: Integer;
    procedure SetValue(v: Integer);
  public
    property Value: Integer read FValue write SetValue;
    property Changes: Integer read FChanges;
  end;

procedure TObservable.SetValue(v: Integer);
begin
  if v <> FValue then begin
    FValue := v;
    FChanges := FChanges + 1;
  end;
end;

var
  o: TObservable;
begin
  o := TObservable.Create;
  o.Value := 1;
  o.Value := 1;
  o.Value := 2;
  o.Value := 3;
  o.Value := 3;
  WriteLn(o.Value);
  WriteLn(o.Changes);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["3", "3"]);
}

#[test]
fn test_property_multiple_in_class() {
    let src = r#"
program T;
type
  TConfig = class
  private
    FHost: string;
    FPort: Integer;
    FDebug: Boolean;
  public
    property Host: string read FHost write FHost;
    property Port: Integer read FPort write FPort;
    property Debug: Boolean read FDebug write FDebug;
    function Summary: string;
  end;

function TConfig.Summary: string;
begin
  Result := FHost + ':' + IntToStr(FPort);
  if FDebug then Result := Result + ' [debug]';
end;

var
  cfg: TConfig;
begin
  cfg := TConfig.Create;
  cfg.Host := 'localhost';
  cfg.Port := 8080;
  cfg.Debug := true;
  WriteLn(cfg.Summary);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["localhost:8080 [debug]"]);
}

#[test]
fn test_property_class_level() {
    let src = r#"
program T;
type
  TApp = class
  private
    class var FVersion: string;
    class function GetVersion: string; static;
    class procedure SetVersion(v: string); static;
  public
    class property Version: string read GetVersion write SetVersion;
  end;

class function TApp.GetVersion: string;
begin
  Result := FVersion;
end;

class procedure TApp.SetVersion(v: string);
begin
  FVersion := v;
end;

begin
  TApp.Version := '1.0.0';
  WriteLn(TApp.Version);
  TApp.Version := '2.0.0';
  WriteLn(TApp.Version);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1.0.0", "2.0.0"]);
}

#[test]
fn test_property_boolean_toggle() {
    let src = r#"
program T;
type
  TSwitch = class
  private
    FOn: Boolean;
    procedure SetOn(v: Boolean);
  public
    property IsOn: Boolean read FOn write SetOn;
    procedure Toggle;
  end;

procedure TSwitch.SetOn(v: Boolean);
begin
  FOn := v;
end;

procedure TSwitch.Toggle;
begin
  FOn := not FOn;
end;

var
  sw: TSwitch;
begin
  sw := TSwitch.Create;
  sw.IsOn := false;
  WriteLn(sw.IsOn);
  sw.Toggle;
  WriteLn(sw.IsOn);
  sw.Toggle;
  WriteLn(sw.IsOn);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["false", "true", "false"]);
}

#[test]
fn test_property_lazy_init() {
    let src = r#"
program T;
type
  TExpensive = class
  private
    FReady: Boolean;
    FData: Integer;
    function GetData: Integer;
  public
    property Data: Integer read GetData;
  end;

function TExpensive.GetData: Integer;
begin
  if not FReady then begin
    FData := 42;
    FReady := true;
  end;
  Result := FData;
end;

var
  e: TExpensive;
begin
  e := TExpensive.Create;
  WriteLn(e.Data);
  WriteLn(e.Data);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["42", "42"]);
}

#[test]
fn test_property_accumulator() {
    let src = r#"
program T;
type
  TAccum = class
  private
    FSum: Integer;
    FCount: Integer;
    function GetAvg: Integer;
  public
    procedure Add(v: Integer);
    property Sum: Integer read FSum;
    property Count: Integer read FCount;
    property Avg: Integer read GetAvg;
  end;

function TAccum.GetAvg: Integer;
begin
  if FCount = 0 then Result := 0
  else Result := FSum div FCount;
end;

procedure TAccum.Add(v: Integer);
begin
  FSum := FSum + v;
  FCount := FCount + 1;
end;

var
  a: TAccum;
begin
  a := TAccum.Create;
  a.Add(10);
  a.Add(20);
  a.Add(30);
  WriteLn(a.Sum);
  WriteLn(a.Count);
  WriteLn(a.Avg);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["60", "3", "20"]);
}

#[test]
fn test_property_protected_in_child() {
    let src = r#"
program T;
type
  TBase = class
  protected
    FName: string;
  public
    property Name: string read FName write FName;
    function Hello: string; virtual;
  end;
  TChild = class(TBase)
    function Hello: string; override;
  end;

function TBase.Hello: string;
begin
  Result := 'Hi, ' + FName;
end;

function TChild.Hello: string;
begin
  Result := 'Hello, ' + FName + '!';
end;

var
  b: TBase;
  c: TChild;
begin
  b := TBase.Create;
  b.Name := 'World';
  WriteLn(b.Hello);
  c := TChild.Create;
  c.Name := 'Pascal';
  WriteLn(c.Hello);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Hi, World", "Hello, Pascal!"]);
}
