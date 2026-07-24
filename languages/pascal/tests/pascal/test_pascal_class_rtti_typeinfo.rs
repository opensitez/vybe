use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 19: RTTI, TypeInfo & Runtime Type Reflection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_classname_reflection() {
    let out = run_pascal(r#"
program Test;
type TCustomWidget = class end;
var obj: TCustomWidget;
begin
  obj := TCustomWidget.Create;
  WriteLn(obj.ClassName);
  obj.Free;
end.
"#);
    assert_eq!(out, vec!["TCustomWidget"]);
}

#[test]
fn test_inheritsfrom_type_check() {
    let out = run_pascal(r#"
program Test;
type TParent = class end;
type TChild = class(TParent) end;
var c: TChild;
begin
  c := TChild.Create;
  WriteLn(c.InheritsFrom(TParent));
  WriteLn(c.InheritsFrom(TObject));
  c.Free;
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_classparent_reflection() {
    let out = run_pascal(r#"
program Test;
type TBaseNode = class end;
type TSubNode = class(TBaseNode) end;
var node: TSubNode;
begin
  node := TSubNode.Create;
  WriteLn(node.ClassParent.ClassName);
  node.Free;
end.
"#);
    assert_eq!(out, vec!["TBaseNode"]);
}

#[test]
fn test_class_reference_type_metaclass() {
    let out = run_pascal(r#"
program Test;
type TAnimal = class
  public constructor Create; virtual;
  public procedure Speak; virtual;
end;
type TDog = class(TAnimal)
  public procedure Speak; override;
end;
type TAnimalClass = class of TAnimal;
constructor TAnimal.Create; begin end;
procedure TAnimal.Speak; begin WriteLn('Animal'); end;
procedure TDog.Speak; begin WriteLn('Woof'); end;
procedure InstantiateAndSpeak(cls: TAnimalClass);
var a: TAnimal;
begin
  a := cls.Create;
  a.Speak;
  a.Free;
end;
begin
  InstantiateAndSpeak(TDog);
end.
"#);
    assert_eq!(out, vec!["Woof"]);
}

#[test]
fn test_classtype_reflection() {
    let out = run_pascal(r#"
program Test;
type TItem = class end;
var item: TItem;
begin
  item := TItem.Create;
  WriteLn(item.ClassType = TItem);
  item.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_typeinfo_typekind_integer() {
    let out = run_pascal(r#"
program Test;
uses Rtti, TypInfo;
begin
  WriteLn(TypeInfo(Integer)^.Kind = tkInteger);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_typeinfo_typekind_string() {
    let out = run_pascal(r#"
program Test;
uses TypInfo;
begin
  WriteLn(TypeInfo(String)^.Kind = tkUString);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_context_type_by_name() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TPerson = class
  public Name: String;
end;
var ctx: TRttiContext;
    t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TPerson);
  WriteLn(t.Name);
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["TPerson"]);
}

#[test]
fn test_rtti_property_reflection() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TCar = class
  private FSpeed: Integer;
  published property Speed: Integer read FSpeed write FSpeed;
end;
var ctx: TRttiContext;
    t: TRttiType;
    p: TRttiProperty;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TCar);
  p := t.GetProperty('Speed');
  WriteLn(p.Name);
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["Speed"]);
}

#[test]
fn test_rtti_read_published_property_value() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TUser = class
  private FAge: Integer;
  public constructor Create(AAge: Integer);
  published property Age: Integer read FAge write FAge;
end;
constructor TUser.Create(AAge: Integer); begin FAge := AAge; end;
var ctx: TRttiContext;
    u: TUser;
    v: TValue;
begin
  u := TUser.Create(35);
  ctx := TRttiContext.Create;
  v := ctx.GetType(TUser).GetProperty('Age').GetValue(u);
  WriteLn(v.AsInteger);
  u.Free;
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["35"]);
}

#[test]
fn test_rtti_write_published_property_value() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TConfig = class
  private FTimeout: Integer;
  published property Timeout: Integer read FTimeout write FTimeout;
end;
var ctx: TRttiContext;
    cfg: TConfig;
begin
  cfg := TConfig.Create;
  ctx := TRttiContext.Create;
  ctx.GetType(TConfig).GetProperty('Timeout').SetValue(cfg, 60);
  WriteLn(cfg.Timeout);
  cfg.Free;
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_rtti_method_reflection_and_invoke() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TCalculator = class
  published function DoubleVal(n: Integer): Integer;
end;
function TCalculator.DoubleVal(n: Integer): Integer; begin Result := n * 2; end;
var ctx: TRttiContext;
    calc: TCalculator;
    m: TRttiMethod;
    res: TValue;
begin
  calc := TCalculator.Create;
  ctx := TRttiContext.Create;
  m := ctx.GetType(TCalculator).GetMethod('DoubleVal');
  res := m.Invoke(calc, [21]);
  WriteLn(res.AsInteger);
  calc.Free;
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_rtti_enum_type_names() {
    let out = run_pascal(r#"
program Test;
uses Rtti, TypInfo;
type TColor = (cRed, cGreen, cBlue);
var info: PTypeInfo;
begin
  info := TypeInfo(TColor);
  WriteLn(GetEnumName(info, Ord(cGreen)));
end.
"#);
    assert_eq!(out, vec!["cGreen"]);
}

#[test]
fn test_rtti_enum_name_to_value() {
    let out = run_pascal(r#"
program Test;
uses TypInfo;
type TStatus = (stInit, stRunning, stFinished);
var info: PTypeInfo;
begin
  info := TypeInfo(TStatus);
  WriteLn(GetEnumValue(info, 'stFinished'));
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_rtti_record_type_reflection() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TPoint = record X, Y: Integer; end;
var ctx: TRttiContext;
    t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TypeInfo(TPoint));
  WriteLn(t.TypeKind = tkRecord);
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_class_fields_enumeration() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TSample = class
  public ID: Integer; Name: String;
end;
var ctx: TRttiContext;
    t: TRttiType;
    fields: TArray<TRttiField>;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TSample);
  fields := t.GetFields;
  WriteLn(Length(fields) >= 2);
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_is_instance_check() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type TBase = class end;
var ctx: TRttiContext;
    t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TBase);
  WriteLn(t.IsInstance);
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_tvalue_from_string() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
var val: TValue;
begin
  val := TValue.FromString('PascalRTTI');
  WriteLn(val.AsString);
end.
"#);
    assert_eq!(out, vec!["PascalRTTI"]);
}

#[test]
fn test_rtti_tvalue_from_integer() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
var val: TValue;
begin
  val := TValue.From<Integer>(500);
  WriteLn(val.AsInteger);
end.
"#);
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_rtti_method_parameter_count() {
    let out = run_pascal(r#"
program Test;
uses Rtti;
type THelper = class
  published procedure DoAction(a, b: Integer);
end;
procedure THelper.DoAction(a, b: Integer); begin end;
var ctx: TRttiContext;
    m: TRttiMethod;
begin
  ctx := TRttiContext.Create;
  m := ctx.GetType(THelper).GetMethod('DoAction');
  WriteLn(Length(m.GetParameters));
  ctx.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}
