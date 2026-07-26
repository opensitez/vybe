use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 11: Class Constructors, Destructors & Lifecycles
// ═══════════════════════════════════════════════════════════

#[test]
fn test_constructor_default_creation() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class
  public constructor Create;
end;
constructor TBase.Create;
begin
  WriteLn('Created');
end;
var obj: TBase;
begin
  obj := TBase.Create;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Created"]);
}

#[test]
fn test_constructor_parameterized_fields() {
    let out = run_pascal(
        r#"
program Test;
type TPerson = class
  public Name: String; Age: Integer;
  constructor Create(AName: String; AAge: Integer);
end;
constructor TPerson.Create(AName: String; AAge: Integer);
begin
  Name := AName; Age := AAge;
end;
var p: TPerson;
begin
  p := TPerson.Create('Alice', 28);
  WriteLn(p.Name + ':' + p.Age.ToString);
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Alice:28"]);
}

#[test]
fn test_destructor_destroy_invocation() {
    let out = run_pascal(
        r#"
program Test;
type TResource = class
  public destructor Destroy; override;
end;
destructor TResource.Destroy;
begin
  WriteLn('Destroyed');
  inherited Destroy;
end;
var r: TResource;
begin
  r := TResource.Create;
  r.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Destroyed"]);
}

#[test]
fn test_free_on_nil_instance_safety() {
    let out = run_pascal(
        r#"
program Test;
type TItem = class end;
var item: TItem;
begin
  item := nil;
  item.Free;
  WriteLn('SafeNilFree');
end.
"#,
    );
    assert_eq!(out, vec!["SafeNilFree"]);
}

#[test]
fn test_inherited_constructor_chaining() {
    let out = run_pascal(
        r#"
program Test;
type TParent = class
  public constructor Create;
end;
type TChild = class(TParent)
  public constructor Create;
end;
constructor TParent.Create;
begin
  WriteLn('ParentCreated');
end;
constructor TChild.Create;
begin
  inherited Create;
  WriteLn('ChildCreated');
end;
var c: TChild;
begin
  c := TChild.Create;
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ParentCreated", "ChildCreated"]);
}

#[test]
fn test_destructor_frees_child_object() {
    let out = run_pascal(
        r#"
program Test;
type TSubObj = class
  public destructor Destroy; override;
end;
type TMainObj = class
  private FSub: TSubObj;
  public constructor Create; destructor Destroy; override;
end;
destructor TSubObj.Destroy;
begin
  WriteLn('SubDestroyed');
  inherited Destroy;
end;
constructor TMainObj.Create;
begin
  FSub := TSubObj.Create;
end;
destructor TMainObj.Destroy;
begin
  FSub.Free;
  WriteLn('MainDestroyed');
  inherited Destroy;
end;
var m: TMainObj;
begin
  m := TMainObj.Create;
  m.Free;
end.
"#,
    );
    assert_eq!(out, vec!["SubDestroyed", "MainDestroyed"]);
}

#[test]
fn test_constructor_overloading() {
    let out = run_pascal(
        r#"
program Test;
type TBox = class
  public Size: Integer;
  constructor Create; overload;
  constructor Create(ASize: Integer); overload;
end;
constructor TBox.Create;
begin
  Size := 10;
end;
constructor TBox.Create(ASize: Integer);
begin
  Size := ASize;
end;
var b1, b2: TBox;
begin
  b1 := TBox.Create;
  b2 := TBox.Create(50);
  WriteLn(b1.Size);
  WriteLn(b2.Size);
  b1.Free; b2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["10", "50"]);
}

#[test]
fn test_constructor_initializes_array_field() {
    let out = run_pascal(
        r#"
program Test;
type TContainer = class
  public Items: array[1..3] of Integer;
  constructor Create;
end;
constructor TContainer.Create;
begin
  Items[1] := 100;
  Items[2] := 200;
  Items[3] := 300;
end;
var c: TContainer;
begin
  c := TContainer.Create;
  WriteLn(c.Items[2]);
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["200"]);
}

#[test]
fn test_class_constructor_execution() {
    let out = run_pascal(
        r#"
program Test;
type TStaticTest = class
  public class var Counter: Integer;
  class constructor Create;
end;
class constructor TStaticTest.Create;
begin
  Counter := 999;
end;
begin
  WriteLn(TStaticTest.Counter);
end.
"#,
    );
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_constructor_virtual_method_dispatch() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class
  public constructor Create; procedure Init; virtual;
end;
type TSub = class(TBase)
  public procedure Init; override;
end;
procedure TBase.Init; begin WriteLn('BaseInit'); end;
procedure TSub.Init; begin WriteLn('SubInit'); end;
constructor TBase.Create;
begin
  Init;
end;
var obj: TSub;
begin
  obj := TSub.Create;
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["SubInit"]);
}

#[test]
fn test_multiple_instance_creations_in_loop() {
    let out = run_pascal(
        r#"
program Test;
type TCounterObj = class
  public ID: Integer;
  constructor Create(AID: Integer);
end;
constructor TCounterObj.Create(AID: Integer);
begin
  ID := AID;
end;
var items: array[1..3] of TCounterObj;
    i: Integer;
begin
  for i := 1 to 3 do
    items[i] := TCounterObj.Create(i * 10);
  for i := 1 to 3 do
    WriteLn(items[i].ID);
  for i := 1 to 3 do
    items[i].Free;
end.
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_destructor_virtual_dispatch() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class
  public destructor Destroy; override;
end;
type TDerived = class(TBase)
  public destructor Destroy; override;
end;
destructor TBase.Destroy;
begin
  WriteLn('BaseDestroy');
  inherited Destroy;
end;
destructor TDerived.Destroy;
begin
  WriteLn('DerivedDestroy');
  inherited Destroy;
end;
var b: TBase;
begin
  b := TDerived.Create;
  b.Free;
end.
"#,
    );
    assert_eq!(out, vec!["DerivedDestroy", "BaseDestroy"]);
}

#[test]
fn test_constructor_setting_string_defaults() {
    let out = run_pascal(
        r#"
program Test;
type TDocument = class
  public Title, Format: String;
  constructor Create;
end;
constructor TDocument.Create;
begin
  Title := 'Untitled';
  Format := 'TXT';
end;
var d: TDocument;
begin
  d := TDocument.Create;
  WriteLn(d.Title + '.' + d.Format);
  d.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Untitled.TXT"]);
}

#[test]
fn test_constructor_with_record_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TConfigRec = record Code: Integer; Name: String; end;
type TNode = class
  public Conf: TConfigRec;
  constructor Create(C: TConfigRec);
end;
constructor TNode.Create(C: TConfigRec);
begin
  Conf := C;
end;
var rec: TConfigRec;
    n: TNode;
begin
  rec.Code := 200; rec.Name := 'OK';
  n := TNode.Create(rec);
  WriteLn(n.Conf.Code);
  WriteLn(n.Conf.Name);
  n.Free;
end.
"#,
    );
    assert_eq!(out, vec!["200", "OK"]);
}

#[test]
fn test_constructor_with_enum_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TMode = (mOff, mOn, mStandby);
type TDevice = class
  public Mode: TMode;
  constructor Create(AMode: TMode = mStandby);
end;
constructor TDevice.Create(AMode: TMode);
begin
  Mode := AMode;
end;
var d1, d2: TDevice;
begin
  d1 := TDevice.Create;
  d2 := TDevice.Create(mOn);
  WriteLn(Ord(d1.Mode));
  WriteLn(Ord(d2.Mode));
  d1.Free; d2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn test_inherited_constructor_with_arguments() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class
  public Val: Integer;
  constructor Create(AVal: Integer);
end;
type TSub = class(TBase)
  public Extra: Integer;
  constructor Create(AVal, AExtra: Integer);
end;
constructor TBase.Create(AVal: Integer);
begin
  Val := AVal;
end;
constructor TSub.Create(AVal, AExtra: Integer);
begin
  inherited Create(AVal);
  Extra := AExtra;
end;
var s: TSub;
begin
  s := TSub.Create(10, 20);
  WriteLn(s.Val + s.Extra);
  s.Free;
end.
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_class_destructor_cleanup() {
    let out = run_pascal(
        r#"
program Test;
type TStaticClean = class
  public class var Active: Boolean;
  class constructor Create;
  class destructor Destroy;
end;
class constructor TStaticClean.Create;
begin
  Active := True;
end;
class destructor TStaticClean.Destroy;
begin
  Active := False;
end;
begin
  WriteLn(TStaticClean.Active);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_constructor_returns_self_reference() {
    let out = run_pascal(
        r#"
program Test;
type TChain = class
  public Count: Integer;
  constructor Create;
  function Add(n: Integer): TChain;
end;
constructor TChain.Create; begin Count := 0; end;
function TChain.Add(n: Integer): TChain;
begin
  Count := Count + n;
  Result := Self;
end;
var c: TChain;
begin
  c := TChain.Create;
  c.Add(5).Add(10).Add(15);
  WriteLn(c.Count);
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_constructor_instantiates_multiple_fields() {
    let out = run_pascal(
        r#"
program Test;
type TComponentA = class public Name: String; constructor Create(N: String); end;
type TComponentB = class public Code: Integer; constructor Create(C: Integer); end;
type TComposite = class
  public A: TComponentA; B: TComponentB;
  constructor Create; destructor Destroy; override;
end;
constructor TComponentA.Create(N: String); begin Name := N; end;
constructor TComponentB.Create(C: Integer); begin Code := C; end;
constructor TComposite.Create;
begin
  A := TComponentA.Create('CompA');
  B := TComponentB.Create(777);
end;
destructor TComposite.Destroy;
begin
  A.Free; B.Free;
  inherited Destroy;
end;
var comp: TComposite;
begin
  comp := TComposite.Create;
  WriteLn(comp.A.Name);
  WriteLn(comp.B.Code);
  comp.Free;
end.
"#,
    );
    assert_eq!(out, vec!["CompA", "777"]);
}

#[test]
fn test_destructor_resets_fields() {
    let out = run_pascal(
        r#"
program Test;
type TStateTracker = class
  public Status: String;
  constructor Create;
  destructor Destroy; override;
end;
constructor TStateTracker.Create; begin Status := 'OPEN'; end;
destructor TStateTracker.Destroy;
begin
  Status := 'CLOSED';
  WriteLn(Status);
  inherited Destroy;
end;
var st: TStateTracker;
begin
  st := TStateTracker.Create;
  st.Free;
end.
"#,
    );
    assert_eq!(out, vec!["CLOSED"]);
}
