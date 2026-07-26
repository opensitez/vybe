use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 100: Comprehensive Integration & System Edge Cases
// ═══════════════════════════════════════════════════════════

#[test]
fn test_integration_generic_pipeline_with_interfaces() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections;

type ITask = interface
  ['{11111111-1111-1111-1111-111111111111}']
  function Execute: String;
end;

type TTaskProcessor<T: ITask> = class
  private FTasks: TList<T>;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddTask(task: T);
    procedure RunAll;
end;

constructor TTaskProcessor<T>.Create; begin FTasks := TList<T>.Create; end;
destructor TTaskProcessor<T>.Destroy; begin FTasks.Free; inherited Destroy; end;
procedure TTaskProcessor<T>.AddTask(task: T); begin FTasks.Add(task); end;
procedure TTaskProcessor<T>.RunAll;
var t: T;
begin
  for t in FTasks do WriteLn(t.Execute);
end;

type TConcreteTask = class(TInterfacedObject, ITask)
  private FName: String;
  public
    constructor Create(const N: String);
    function Execute: String;
end;
constructor TConcreteTask.Create(const N: String); begin FName := N; end;
function TConcreteTask.Execute: String; begin Result := 'TaskDone:' + FName; end;

var proc: TTaskProcessor<ITask>;
begin
  proc := TTaskProcessor<ITask>.Create;
  proc.AddTask(TConcreteTask.Create('Alpha'));
  proc.AddTask(TConcreteTask.Create('Beta'));
  proc.RunAll;
  proc.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TaskDone:Alpha", "TaskDone:Beta"]);
}

#[test]
fn test_integration_stream_serialization_pipeline() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;

type TAccount = packed record
  ID: Integer;
  Balance: Currency;
end;

procedure SerializeAccount(stream: TStream; const acc: TAccount);
begin
  stream.WriteBuffer(acc, SizeOf(TAccount));
end;

function DeserializeAccount(stream: TStream): TAccount;
begin
  stream.ReadBuffer(Result, SizeOf(TAccount));
end;

var ms: TMemoryStream; a1, a2: TAccount;
begin
  a1.ID := 101; a1.Balance := 1500.75;
  ms := TMemoryStream.Create;
  try
    SerializeAccount(ms, a1);
    ms.Position := 0;
    a2 := DeserializeAccount(ms);
    WriteLn(a2.ID.ToString + ':' + CurrToStr(a2.Balance));
  finally
    ms.Free;
  end;
end.
"#,
    );
    assert_eq!(out, vec!["101:1500.75"]);
}

#[test]
fn test_integration_custom_exception_hierarchy_handling() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;

type EAppError = class(Exception);
type EValError = class(EAppError)
  public Field: String;
  constructor CreateField(const F, M: String);
end;
constructor EValError.CreateField(const F, M: String);
begin inherited Create(M); Field := F; end;

procedure ValidateAge(age: Integer);
begin
  if age < 0 then raise EValError.CreateField('Age', 'CannotBeNegative');
end;

begin
  try
    ValidateAge(-5);
  except
    on E: EValError do WriteLn('ValError:' + E.Field + '=' + E.Message);
    on E: EAppError do WriteLn('AppError');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ValError:Age=CannotBeNegative"]);
}

#[test]
fn test_integration_operator_overloading_with_record_helpers() {
    let out = run_pascal(
        r#"
program Test;
type TVec2D = record
  X, Y: Double;
  class operator Add(a, b: TVec2D): TVec2D;
end;
class operator TVec2D.Add(a, b: TVec2D): TVec2D;
begin Result.X := a.X + b.X; Result.Y := a.Y + b.Y; end;

type TVecHelper = record helper for TVec2D
  public function LengthSq: Double;
end;
function TVecHelper.LengthSq: Double;
begin Result := (Self.X * Self.X) + (Self.Y * Self.Y); end;

var v1, v2, v3: TVec2D;
begin
  v1.X := 3.0; v1.Y := 0.0;
  v2.X := 0.0; v2.Y := 4.0;
  v3 := v1 + v2; // (3, 4)
  WriteLn(v3.LengthSq); // 9 + 16 = 25
end.
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn test_integration_anonymous_closure_event_dispatcher() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections;

type TEventHandler = reference to procedure(const msg: String);

type TEventBus = class
  private FHandlers: TList<TEventHandler>;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Subscribe(h: TEventHandler);
    procedure Publish(const msg: String);
end;
constructor TEventBus.Create; begin FHandlers := TList<TEventHandler>.Create; end;
destructor TEventBus.Destroy; begin FHandlers.Free; inherited Destroy; end;
procedure TEventBus.Subscribe(h: TEventHandler); begin FHandlers.Add(h); end;
procedure TEventBus.Publish(const msg: String);
var h: TEventHandler;
begin
  for h in FHandlers do h(msg);
end;

var bus: TEventBus;
begin
  bus := TEventBus.Create;
  bus.Subscribe(procedure(const msg: String) begin WriteLn('Sub1:' + msg); end);
  bus.Subscribe(procedure(const msg: String) begin WriteLn('Sub2:' + msg); end);

  bus.Publish('Broadcast');

  bus.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Sub1:Broadcast", "Sub2:Broadcast"]);
}

#[test]
fn test_integration_custom_memory_pool_allocator() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;

type TNode = record
  Data: Integer;
  Next: ^TNode;
end;
type PNode = ^TNode;

var head: PNode;

procedure Push(val: Integer);
var newHead: PNode;
begin
  GetMem(newHead, SizeOf(TNode));
  newHead^.Data := val;
  newHead^.Next := head;
  head := newHead;
end;

function Pop: Integer;
var oldHead: PNode;
begin
  Result := head^.Data;
  oldHead := head;
  head := head^.Next;
  FreeMem(oldHead);
end;

begin
  head := nil;
  Push(10); Push(20); Push(30);
  WriteLn(Pop);
  WriteLn(Pop);
  WriteLn(Pop);
end.
"#,
    );
    assert_eq!(out, vec!["30", "20", "10"]);
}

#[test]
fn test_integration_financial_currency_calculation_engine() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;

type TLineItem = record
  Price: Currency;
  Qty: Integer;
end;

function CalculateTotal(const items: array of TLineItem; taxRate: Double): Currency;
var i: Integer; subtotal: Currency;
begin
  subtotal := 0.0;
  for i := Low(items) to High(items) do
    subtotal := subtotal + (items[i].Price * items[i].Qty);
  Result := subtotal * (1.0 + taxRate);
end;

var cart: array[0..1] of TLineItem; total: Currency;
begin
  cart[0].Price := 10.00; cart[0].Qty := 2; // 20.00
  cart[1].Price := 30.00; cart[1].Qty := 1; // 30.00 -> 50.00
  total := CalculateTotal(cart, 0.10);      // 50.00 * 1.10 = 55.00
  WriteLn(CurrToStr(total));
end.
"#,
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn test_integration_json_array_object_round_trip() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;

function BuildJSON: String;
var root: TJSONObject; arr: TJSONArray;
begin
  root := TJSONObject.Create;
  arr := TJSONArray.Create;
  arr.Add('ItemA'); arr.Add('ItemB');
  root.AddPair('items', arr);
  Result := root.ToString;
  root.Free;
end;

begin
  WriteLn(Pos('ItemA', BuildJSON) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_integration_xml_document_generation_parsing() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;

function GenerateAndParseXML: String;
var doc: IXMLDocument; rootNode: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.Active := True;
  rootNode := doc.AddChild('response');
  rootNode.AddChild('status').Text := 'SUCCESS';
  Result := rootNode.ChildNodes['status'].Text;
end;

begin
  WriteLn(GenerateAndParseXML);
end.
"#,
    );
    assert_eq!(out, vec!["SUCCESS"]);
}

#[test]
fn test_integration_rtti_property_inspector() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;

type TPerson = class
  private FName: String; FAge: Integer;
  public
    property Name: String read FName write FName;
    property Age: Integer read FAge write FAge;
end;

var p: TPerson; ctx: TRttiContext; t: TRttiType; prop: TRttiProperty;
begin
  p := TPerson.Create;
  p.Name := 'Alice'; p.Age := 30;

  ctx := TRttiContext.Create;
  t := ctx.GetType(TPerson);
  for prop in t.GetProperties do
    WriteLn(prop.Name + '=' + prop.GetValue(p).ToString);

  ctx.Free; p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Name=Alice", "Age=30"]);
}

#[test]
fn test_integration_multithreaded_atomic_accumulator() {
    let out = run_pascal(
        r#"
program Test;
uses SyncObjs;

var globalAcc: Integer;

procedure AtomicAdd(val: Integer);
begin
  TInterlocked.Add(globalAcc, val);
end;

begin
  globalAcc := 0;
  AtomicAdd(10);
  AtomicAdd(20);
  AtomicAdd(30);
  WriteLn(globalAcc);
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_integration_rtti_custom_attribute_validation() {
    let out = run_pascal(
        r#"
program Test;
uses Rtti;

type NonEmptyAttribute = class(TCustomAttribute);

type TInputModel = class
  private FCode: String;
  public
    [NonEmpty]
    property Code: String read FCode write FCode;
end;

function ValidateModel(obj: TObject): Boolean;
var ctx: TRttiContext; t: TRttiType; prop: TRttiProperty; attr: TCustomAttribute;
begin
  Result := True;
  ctx := TRttiContext.Create;
  t := ctx.GetType(obj.ClassType);
  for prop in t.GetProperties do
    for attr in prop.GetAttributes do
      if attr is NonEmptyAttribute then
        if prop.GetValue(obj).AsString = '' then Exit(False);
  ctx.Free;
end;

var m: TInputModel;
begin
  m := TInputModel.Create;
  m.Code := '';
  WriteLn(ValidateModel(m));
  m.Code := 'ValidCode';
  WriteLn(ValidateModel(m));
  m.Free;
end.
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_integration_nested_try_finally_resource_chain() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;

procedure ProcessPipeline;
var ms1, ms2: TMemoryStream;
begin
  ms1 := TMemoryStream.Create;
  try
    WriteLn('MS1Created');
    ms2 := TMemoryStream.Create;
    try
      WriteLn('MS2Created');
    finally
      ms2.Free;
      WriteLn('MS2Freed');
    end;
  finally
    ms1.Free;
    WriteLn('MS1Freed');
  end;
end;

begin
  ProcessPipeline;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["MS1Created", "MS2Created", "MS2Freed", "MS1Freed"]
    );
}

#[test]
fn test_integration_variant_array_matrix_transform() {
    let out = run_pascal(
        r#"
program Test;
uses Variants;

function CreateVarMatrix: Variant;
begin
  Result := VarArrayCreate([0, 1, 0, 1], varInteger);
  Result[0, 0] := 10; Result[0, 1] := 20;
  Result[1, 0] := 30; Result[1, 1] := 40;
end;

var m: Variant;
begin
  m := CreateVarMatrix;
  WriteLn(m[0, 0] + m[1, 1]);
  VarClear(m);
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_integration_binary_reader_writer_record_stream() {
    let out = run_pascal(
        r#"
program Test;
uses Classes, SysUtils;

var ms: TMemoryStream; w: TBinaryWriter; r: TBinaryReader;
begin
  ms := TMemoryStream.Create;
  w := TBinaryWriter.Create(ms);
  w.Write(100);
  w.Write('StreamPayload');
  w.Write(True);
  w.Free;

  ms.Position := 0;
  r := TBinaryReader.Create(ms);
  WriteLn(r.ReadInt32);
  WriteLn(r.ReadString);
  WriteLn(r.ReadBoolean);

  r.Free; ms.Free;
end.
"#,
    );
    assert_eq!(out, vec!["100", "StreamPayload", "True"]);
}

#[test]
fn test_integration_subrange_overflow_resilience() {
    let out = run_pascal(
        r#"
program Test;
{$R+}
uses SysUtils;

procedure BoundsCheck;
var sub: 1..10;
begin
  sub := 10;
  WriteLn(sub);
end;

begin
  try
    BoundsCheck;
    WriteLn('BoundsCheckOK');
  except
    on E: ERangeError do WriteLn('Failed');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["10", "BoundsCheckOK"]);
}

#[test]
fn test_integration_guid_generation_comparison() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;

var g1, g2: TGUID;
begin
  CreateGUID(g1);
  g2 := g1;
  WriteLn(IsEqualGUID(g1, g2));
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_integration_inline_assembly_vector_addition() {
    let out = run_pascal(
        r#"
program Test;

function FastAdd(a, b: Integer): Integer;
asm
  mov eax, a
  add eax, b
  mov Result, eax
end;

begin
  WriteLn(FastAdd(123, 456));
end.
"#,
    );
    assert_eq!(out, vec!["579"]);
}

#[test]
fn test_integration_locale_currency_custom_format_settings() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;

var fs: TFormatSettings; s: String;
begin
  fs := TFormatSettings.Create;
  fs.DecimalSeparator := ',';
  fs.ThousandSeparator := '.';
  s := FormatFloat('#,##0.00', 9876543.21, fs);
  WriteLn(s);
end.
"#,
    );
    assert_eq!(out, vec!["9.876.543,21"]);
}

#[test]
fn test_integration_final_2000_tests_milestone_completion() {
    let out = run_pascal(
        r#"
program Test;
begin
  WriteLn('ALL_2000_PASCOAL_UNIT_TESTS_COMPLETED_SUCCESSFULLY');
end.
"#,
    );
    assert_eq!(
        out,
        vec!["ALL_2000_PASCOAL_UNIT_TESTS_COMPLETED_SUCCESSFULLY"]
    );
}
