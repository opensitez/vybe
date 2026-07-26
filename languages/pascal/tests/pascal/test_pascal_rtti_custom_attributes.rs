use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 81: RTTI Custom Attributes & Metadata Annotations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_rtti_custom_attribute_class_declaration() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type TableAttribute = class(TCustomAttribute)
  private FTableName: String;
  public constructor Create(const AName: String);
  public property TableName: String read FTableName;
end;
constructor TableAttribute.Create(const AName: String);
begin
  FTableName := AName;
end;

type
  [Table('users_table')]
  TUserObj = class end;

var ctx: TRttiContext; t: TRttiType; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TUserObj);
  for attr in t.GetAttributes do
    if attr is TableAttribute then
      WriteLn(TableAttribute(attr).TableName);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["users_table"]);
}

#[test]
fn test_rtti_custom_attribute_on_property() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type RequiredAttribute = class(TCustomAttribute);
type TFormModel = class
  private FEmail: String;
  public
    [Required]
    property Email: String read FEmail write FEmail;
end;

var ctx: TRttiContext; t: TRttiType; p: TRttiProperty; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TFormModel);
  p := t.GetProperty('Email');
  for attr in p.GetAttributes do
    if attr is RequiredAttribute then
      WriteLn('EmailIsRequired');
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["EmailIsRequired"]);
}

#[test]
fn test_rtti_custom_attribute_on_method() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type RouteAttribute = class(TCustomAttribute)
  public Path: String;
  constructor Create(const APath: String);
end;
constructor RouteAttribute.Create(const APath: String); begin Path := APath; end;

type TApiController = class
  public
    [Route('/api/v1/status')]
    procedure GetStatus;
end;
procedure TApiController.GetStatus; begin end;

var ctx: TRttiContext; t: TRttiType; m: TRttiMethod; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TApiController);
  m := t.GetMethod('GetStatus');
  for attr in m.GetAttributes do
    if attr is RouteAttribute then
      WriteLn(RouteAttribute(attr).Path);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["/api/v1/status"]);
}

#[test]
fn test_rtti_multiple_attributes_on_single_symbol() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type AttrA = class(TCustomAttribute);
type AttrB = class(TCustomAttribute);

type
  [AttrA]
  [AttrB]
  TMultiObj = class end;

var ctx: TRttiContext; t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TMultiObj);
  WriteLn(Length(t.GetAttributes) = 2);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_attribute_integer_arg() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type MaxLengthAttribute = class(TCustomAttribute)
  public MaxLen: Integer;
  constructor Create(L: Integer);
end;
constructor MaxLengthAttribute.Create(L: Integer); begin MaxLen := L; end;

type TDTO = class
  private FName: String;
  public
    [MaxLength(50)]
    property Name: String read FName write FName;
end;

var ctx: TRttiContext; t: TRttiType; p: TRttiProperty; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TDTO);
  p := t.GetProperty('Name');
  for attr in p.GetAttributes do
    if attr is MaxLengthAttribute then
      WriteLn(MaxLengthAttribute(attr).MaxLen);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_rtti_attribute_inheritance() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type BaseAttr = class(TCustomAttribute);
type DerivedAttr = class(BaseAttr);

type
  [DerivedAttr]
  TAnnotated = class end;

var ctx: TRttiContext; t: TRttiType; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TAnnotated);
  for attr in t.GetAttributes do
    WriteLn(attr is BaseAttr);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_custom_attribute_on_field() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type FieldTagAttribute = class(TCustomAttribute)
  public Tag: String;
  constructor Create(const T: String);
end;
constructor FieldTagAttribute.Create(const T: String); begin Tag := T; end;

type TRecordFieldObj = class
  public
    [FieldTag('primary_key')]
    FID: Integer;
end;

var ctx: TRttiContext; t: TRttiType; f: TRttiField; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TRecordFieldObj);
  f := t.GetField('FID');
  for attr in f.GetAttributes do
    if attr is FieldTagAttribute then
      WriteLn(FieldTagAttribute(attr).Tag);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["primary_key"]);
}

#[test]
fn test_rtti_custom_attribute_enum_argument() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type TSeverity = (sevLow, sevHigh);
type SeverityAttribute = class(TCustomAttribute)
  public Level: TSeverity;
  constructor Create(L: TSeverity);
end;
constructor SeverityAttribute.Create(L: TSeverity); begin Level := L; end;

type
  [Severity(sevHigh)]
  TCriticalComponent = class end;

var ctx: TRttiContext; t: TRttiType; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TCriticalComponent);
  for attr in t.GetAttributes do
    if attr is SeverityAttribute then
      WriteLn(Ord(SeverityAttribute(attr).Level));
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_rtti_custom_attribute_boolean_argument() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type VisibleAttribute = class(TCustomAttribute)
  public IsVisible: Boolean;
  constructor Create(V: Boolean);
end;
constructor VisibleAttribute.Create(V: Boolean); begin IsVisible := V; end;

type
  [Visible(False)]
  THiddenObj = class end;

var ctx: TRttiContext; t: TRttiType; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(THiddenObj);
  for attr in t.GetAttributes do
    if attr is VisibleAttribute then
      WriteLn(VisibleAttribute(attr).IsVisible);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_rtti_custom_attribute_default_constructor() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type SerializableAttribute = class(TCustomAttribute);

type
  [Serializable]
  TPayload = class end;

var ctx: TRttiContext; t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TPayload);
  WriteLn(t.GetAttributes[0].ClassName);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["SerializableAttribute"]);
}

#[test]
fn test_rtti_custom_attribute_on_record() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type RecMetaAttribute = class(TCustomAttribute)
  public Meta: String;
  constructor Create(const M: String);
end;
constructor RecMetaAttribute.Create(const M: String); begin Meta := M; end;

type
  [RecMeta('struct_v1')]
  TMyAnnotatedRec = record
    Val: Integer;
  end;

var ctx: TRttiContext; t: TRttiType; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TypeInfo(TMyAnnotatedRec));
  for attr in t.GetAttributes do
    if attr is RecMetaAttribute then
      WriteLn(RecMetaAttribute(attr).Meta);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["struct_v1"]);
}

#[test]
fn test_rtti_custom_attribute_multiple_params() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type ColumnAttribute = class(TCustomAttribute)
  public Name: String; IsNullable: Boolean;
  constructor Create(const AName: String; Nullable: Boolean);
end;
constructor ColumnAttribute.Create(const AName: String; Nullable: Boolean);
begin
  Name := AName; IsNullable := Nullable;
end;

type TEntity = class
  private FAge: Integer;
  public
    [Column('user_age', True)]
    property Age: Integer read FAge write FAge;
end;

var ctx: TRttiContext; t: TRttiType; p: TRttiProperty; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TEntity);
  p := t.GetProperty('Age');
  for attr in p.GetAttributes do
    if attr is ColumnAttribute then
      WriteLn(ColumnAttribute(attr).Name + ':' + ColumnAttribute(attr).IsNullable.ToString);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["user_age:True"]);
}

#[test]
fn test_rtti_no_attributes_returns_empty_array() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type TPlainClass = class end;
var ctx: TRttiContext; t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TPlainClass);
  WriteLn(Length(t.GetAttributes) = 0);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_attribute_query_by_class() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type TargetAttr = class(TCustomAttribute);
type
  [TargetAttr]
  TSample = class end;

function HasAttr(t: TRttiType; attrClass: TClass): Boolean;
var a: TCustomAttribute;
begin
  Result := False;
  for a in t.GetAttributes do
    if a.InheritsFrom(attrClass) then Exit(True);
end;

var ctx: TRttiContext; t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TSample);
  WriteLn(HasAttr(t, TargetAttr));
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_attribute_on_interface() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type ServiceAttribute = class(TCustomAttribute);

type
  [Service]
  IServiceContract = interface
    ['{12345678-1234-1234-1234-123456789012}']
  end;

var ctx: TRttiContext; t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TypeInfo(IServiceContract));
  WriteLn(Length(t.GetAttributes) = 1);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_attribute_float_argument() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type VersionAttr = class(TCustomAttribute)
  public Ver: Double;
  constructor Create(V: Double);
end;
constructor VersionAttr.Create(V: Double); begin Ver := V; end;

type
  [VersionAttr(2.5)]
  TModule = class end;

var ctx: TRttiContext; t: TRttiType; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TModule);
  for attr in t.GetAttributes do
    if attr is VersionAttr then
      WriteLn(VersionAttr(attr).Ver);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn test_rtti_attribute_on_method_parameter() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type ParamCheckAttr = class(TCustomAttribute);
type TProcessor = class
  public procedure Process([ParamCheckAttr] const val: String);
end;
procedure TProcessor.Process(const val: String); begin end;

var ctx: TRttiContext; t: TRttiType; m: TRttiMethod; params: TArray<TRttiParameter>;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TProcessor);
  m := t.GetMethod('Process');
  params := m.GetParameters;
  WriteLn(Length(params[0].GetAttributes) = 1);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_attribute_subclass_polymorphic_filtering() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type BaseValidationAttr = class(TCustomAttribute);
type NotEmptyAttr = class(BaseValidationAttr);

type TForm = class
  private FCode: String;
  public
    [NotEmptyAttr]
    property Code: String read FCode write FCode;
end;

var ctx: TRttiContext; t: TRttiType; p: TRttiProperty; attr: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TForm);
  p := t.GetProperty('Code');
  for attr in p.GetAttributes do
    if attr is BaseValidationAttr then
      WriteLn(attr.ClassName);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["NotEmptyAttr"]);
}

#[test]
fn test_rtti_attribute_instance_uniqueness() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type UniqueAttr = class(TCustomAttribute);

type
  [UniqueAttr]
  TClass1 = class end;
type
  [UniqueAttr]
  TClass2 = class end;

var ctx: TRttiContext; t1, t2: TRttiType; a1, a2: TCustomAttribute;
begin
  ctx := TRttiContext.Create;
  t1 := ctx.GetType(TClass1);
  t2 := ctx.GetType(TClass2);
  a1 := t1.GetAttributes[0];
  a2 := t2.GetAttributes[0];
  WriteLn(a1 <> a2);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_rtti_attribute_on_enum_type() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;
type EnumMetaAttr = class(TCustomAttribute);

type
  [EnumMetaAttr]
  TColorEnum = (cRed, cGreen, cBlue);

var ctx: TRttiContext; t: TRttiType;
begin
  ctx := TRttiContext.Create;
  t := ctx.GetType(TypeInfo(TColorEnum));
  WriteLn(Length(t.GetAttributes) = 1);
  ctx.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}
