use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 20: Class Visibility Specifiers (private, protected, public, published, strict)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_private_field_encapsulation_access_via_public_method() {
    let out = run_pascal(r#"
program Test;
type TAccount = class
  private FBalance: Integer;
  public constructor Create(initial: Integer);
  public function GetBalance: Integer;
end;
constructor TAccount.Create(initial: Integer); begin FBalance := initial; end;
function TAccount.GetBalance: Integer; begin Result := FBalance; end;
var acc: TAccount;
begin
  acc := TAccount.Create(500);
  WriteLn(acc.GetBalance);
  acc.Free;
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_protected_method_accessed_by_subclass() {
    let out = run_pascal(r#"
program Test;
type TBaseProcessor = class
  protected procedure InternalProcess; virtual;
end;
type TCustomProcessor = class(TBaseProcessor)
  public procedure Execute;
  protected procedure InternalProcess; override;
end;
procedure TBaseProcessor.InternalProcess; begin WriteLn('BaseInternal'); end;
procedure TCustomProcessor.InternalProcess; begin WriteLn('CustomInternal'); end;
procedure TCustomProcessor.Execute;
begin
  InternalProcess;
end;
var p: TCustomProcessor;
begin
  p := TCustomProcessor.Create;
  p.Execute;
  p.Free;
end.
"#);
    assert_eq!(out, vec!["CustomInternal"]);
}

#[test]
fn test_public_field_direct_access() {
    let out = run_pascal(r#"
program Test;
type TPointObj = class
  public X, Y: Integer;
end;
var pt: TPointObj;
begin
  pt := TPointObj.Create;
  pt.X := 10;
  pt.Y := 20;
  WriteLn(pt.X + pt.Y);
  pt.Free;
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_published_property_accessibility() {
    let out = run_pascal(r#"
program Test;
type TWidget = class
  private FTitle: String;
  published property Title: String read FTitle write FTitle;
end;
var w: TWidget;
begin
  w := TWidget.Create;
  w.Title := 'MainWidget';
  WriteLn(w.Title);
  w.Free;
end.
"#);
    assert_eq!(out, vec!["MainWidget"]);
}

#[test]
fn test_strict_private_field_isolation() {
    let out = run_pascal(r#"
program Test;
type TVault = class
  strict private FSecretCode: Integer;
  public constructor Create(code: Integer);
  public function Validate(input: Integer): Boolean;
end;
constructor TVault.Create(code: Integer); begin FSecretCode := code; end;
function TVault.Validate(input: Integer): Boolean; begin Result := FSecretCode = input; end;
var v: TVault;
begin
  v := TVault.Create(1234);
  WriteLn(v.Validate(1234));
  WriteLn(v.Validate(9999));
  v.Free;
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_strict_protected_field_accessible_to_subclass() {
    let out = run_pascal(r#"
program Test;
type TBaseData = class
  strict protected FData: String;
end;
type TSubData = class(TBaseData)
  public procedure SetData(s: String);
  public function GetData: String;
end;
procedure TSubData.SetData(s: String); begin FData := s; end;
function TSubData.GetData: String; begin Result := FData; end;
var sd: TSubData;
begin
  sd := TSubData.Create;
  sd.SetData('StrictProtectedValue');
  WriteLn(sd.GetData);
  sd.Free;
end.
"#);
    assert_eq!(out, vec!["StrictProtectedValue"]);
}

#[test]
fn test_same_class_instances_accessing_private_fields() {
    let out = run_pascal(r#"
program Test;
type TCompareObj = class
  private FVal: Integer;
  public constructor Create(V: Integer);
  public function IsEqual(other: TCompareObj): Boolean;
end;
constructor TCompareObj.Create(V: Integer); begin FVal := V; end;
function TCompareObj.IsEqual(other: TCompareObj): Boolean;
begin
  Result := Self.FVal = other.FVal;
end;
var o1, o2: TCompareObj;
begin
  o1 := TCompareObj.Create(10);
  o2 := TCompareObj.Create(10);
  WriteLn(o1.IsEqual(o2));
  o1.Free; o2.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_private_method_called_by_public_method() {
    let out = run_pascal(r#"
program Test;
type TPipeline = class
  private procedure Step1;
  private procedure Step2;
  public procedure RunAll;
end;
procedure TPipeline.Step1; begin WriteLn('Step1'); end;
procedure TPipeline.Step2; begin WriteLn('Step2'); end;
procedure TPipeline.RunAll;
begin
  Step1;
  Step2;
end;
var p: TPipeline;
begin
  p := TPipeline.Create;
  p.RunAll;
  p.Free;
end.
"#);
    assert_eq!(out, vec!["Step1", "Step2"]);
}

#[test]
fn test_protected_field_mutation_in_derived_constructor() {
    let out = run_pascal(r#"
program Test;
type TBaseEntity = class
  protected FID: Integer;
end;
type TUserEntity = class(TBaseEntity)
  public constructor Create(AID: Integer);
  public function GetID: Integer;
end;
constructor TUserEntity.Create(AID: Integer); begin FID := AID; end;
function TUserEntity.GetID: Integer; begin Result := FID; end;
var u: TUserEntity;
begin
  u := TUserEntity.Create(777);
  WriteLn(u.GetID);
  u.Free;
end.
"#);
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_public_property_wrapping_private_setter() {
    let out = run_pascal(r#"
program Test;
type TScoreTracker = class
  private FScore: Integer;
  private procedure SetScore(v: Integer);
  public property Score: Integer read FScore write SetScore;
end;
procedure TScoreTracker.SetScore(v: Integer);
begin
  if v >= 0 then FScore := v;
end;
var st: TScoreTracker;
begin
  st := TScoreTracker.Create;
  st.Score := 50;
  WriteLn(st.Score);
  st.Free;
end.
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_multiple_visibility_sections_in_single_class() {
    let out = run_pascal(r#"
program Test;
type TComplexClass = class
  private FPriv: Integer;
  protected FProt: Integer;
  public FPub: Integer;
  public constructor Create;
  public function GetSum: Integer;
end;
constructor TComplexClass.Create; begin FPriv := 1; FProt := 2; FPub := 3; end;
function TComplexClass.GetSum: Integer; begin Result := FPriv + FProt + FPub; end;
var c: TComplexClass;
begin
  c := TComplexClass.Create;
  WriteLn(c.GetSum);
  c.Free;
end.
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_protected_virtual_method_override_access() {
    let out = run_pascal(r#"
program Test;
type TBaseApp = class
  protected function GetName: String; virtual;
  public procedure PrintName;
end;
type TCustomApp = class(TBaseApp)
  protected function GetName: String; override;
end;
function TBaseApp.GetName: String; begin Result := 'BaseApp'; end;
procedure TBaseApp.PrintName; begin WriteLn(GetName); end;
function TCustomApp.GetName: String; begin Result := 'CustomApp'; end;
var app: TBaseApp;
begin
  app := TCustomApp.Create;
  app.PrintName;
  app.Free;
end.
"#);
    assert_eq!(out, vec!["CustomApp"]);
}

#[test]
fn test_public_method_exposing_private_record() {
    let out = run_pascal(r#"
program Test;
type TConfigData = record Key, Val: String; end;
type TConfigStore = class
  private FData: TConfigData;
  public constructor Create(K, V: String);
  public function GetData: TConfigData;
end;
constructor TConfigStore.Create(K, V: String); begin FData.Key := K; FData.Val := V; end;
function TConfigStore.GetData: TConfigData; begin Result := FData; end;
var cs: TConfigStore; d: TConfigData;
begin
  cs := TConfigStore.Create('env', 'production');
  d := cs.GetData;
  WriteLn(d.Key + '=' + d.Val);
  cs.Free;
end.
"#);
    assert_eq!(out, vec!["env=production"]);
}

#[test]
fn test_strict_private_method_call_internal() {
    let out = run_pascal(r#"
program Test;
type TSecurityCheck = class
  strict private function InternalCheck(code: Integer): Boolean;
  public function Validate(code: Integer): Boolean;
end;
function TSecurityCheck.InternalCheck(code: Integer): Boolean; begin Result := code = 42; end;
function TSecurityCheck.Validate(code: Integer): Boolean; begin Result := InternalCheck(code); end;
var sc: TSecurityCheck;
begin
  sc := TSecurityCheck.Create;
  WriteLn(sc.Validate(42));
  sc.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_published_property_backed_by_private_getter() {
    let out = run_pascal(r#"
program Test;
type TSettings = class
  private FPort: Integer;
  private function GetPort: Integer;
  public constructor Create;
  published property Port: Integer read GetPort;
end;
constructor TSettings.Create; begin FPort := 8080; end;
function TSettings.GetPort: Integer; begin Result := FPort; end;
var s: TSettings;
begin
  s := TSettings.Create;
  WriteLn(s.Port);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["8080"]);
}

#[test]
fn test_protected_method_with_default_parameters() {
    let out = run_pascal(r#"
program Test;
type TBaseWorker = class
  protected procedure Work(intensity: Integer = 1); virtual;
  public procedure Execute;
end;
type TFastWorker = class(TBaseWorker)
  protected procedure Work(intensity: Integer = 1); override;
end;
procedure TBaseWorker.Work(intensity: Integer); begin WriteLn('BaseWork:' + intensity.ToString); end;
procedure TBaseWorker.Execute; begin Work; end;
procedure TFastWorker.Work(intensity: Integer); begin WriteLn('FastWork:' + intensity.ToString); end;
var w: TBaseWorker;
begin
  w := TFastWorker.Create;
  w.Execute;
  w.Free;
end.
"#);
    assert_eq!(out, vec!["FastWork:1"]);
}

#[test]
fn test_public_class_function_accessing_private_class_var() {
    let out = run_pascal(r#"
program Test;
type TGlobalState = class
  private class var FCount: Integer;
  public class procedure IncCount;
  public class function GetCount: Integer;
end;
class procedure TGlobalState.IncCount; begin Inc(FCount); end;
class function TGlobalState.GetCount: Integer; begin Result := FCount; end;
begin
  TGlobalState.IncCount;
  TGlobalState.IncCount;
  WriteLn(TGlobalState.GetCount);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_private_field_array_accessed_via_public_method() {
    let out = run_pascal(r#"
program Test;
type TArrayStorage = class
  private FElements: array[0..2] of Integer;
  public constructor Create;
  public function GetSum: Integer;
end;
constructor TArrayStorage.Create; begin FElements[0] := 10; FElements[1] := 20; FElements[2] := 30; end;
function TArrayStorage.GetSum: Integer; begin Result := FElements[0] + FElements[1] + FElements[2]; end;
var store: TArrayStorage;
begin
  store := TArrayStorage.Create;
  WriteLn(store.GetSum);
  store.Free;
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_protected_property_promoted_to_public_in_derived_class() {
    let out = run_pascal(r#"
program Test;
type TBaseComponent = class
  protected FTag: Integer;
  protected property Tag: Integer read FTag write FTag;
end;
type TPublicComponent = class(TBaseComponent)
  public property Tag: Integer read FTag write FTag;
end;
var pc: TPublicComponent;
begin
  pc := TPublicComponent.Create;
  pc.Tag := 999;
  WriteLn(pc.Tag);
  pc.Free;
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_private_constructor_pattern() {
    let out = run_pascal(r#"
program Test;
type TSingleton = class
  private constructor Create;
  public class function GetInstance: TSingleton;
  public procedure Speak;
end;
constructor TSingleton.Create; begin end;
class function TSingleton.GetInstance: TSingleton;
begin
  Result := TSingleton.Create;
end;
procedure TSingleton.Speak; begin WriteLn('SingletonActive'); end;
var s: TSingleton;
begin
  s := TSingleton.GetInstance;
  s.Speak;
  s.Free;
end.
"#);
    assert_eq!(out, vec!["SingletonActive"]);
}
