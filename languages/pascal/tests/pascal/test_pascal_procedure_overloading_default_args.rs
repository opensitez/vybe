use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 98: Overloaded Routines & Default Parameter Resolution
// ═══════════════════════════════════════════════════════════

#[test]
fn test_overload_by_parameter_type() {
    let out = run_pascal(r#"
program Test;
procedure PrintVal(v: Integer); overload;
begin WriteLn('Int:' + v.ToString); end;

procedure PrintVal(v: Double); overload;
begin WriteLn('Float:' + FloatToStr(v)); end;

procedure PrintVal(const v: String); overload;
begin WriteLn('Str:' + v); end;

begin
  PrintVal(42);
  PrintVal(3.14);
  PrintVal('Hello');
end.
"#);
    assert_eq!(out, vec!["Int:42", "Float:3.14", "Str:Hello"]);
}

#[test]
fn test_overload_by_parameter_count() {
    let out = run_pascal(r#"
program Test;
function Add(a, b: Integer): Integer; overload;
begin Result := a + b; end;

function Add(a, b, c: Integer): Integer; overload;
begin Result := a + b + c; end;

begin
  WriteLn(Add(10, 20));
  WriteLn(Add(10, 20, 30));
end.
"#);
    assert_eq!(out, vec!["30", "60"]);
}

#[test]
fn test_default_parameter_single() {
    let out = run_pascal(r#"
program Test;
procedure Greet(name: String; prefix: String = 'Hello ');
begin
  WriteLn(prefix + name);
end;
begin
  Greet('Alice');
  Greet('Bob', 'Welcome ');
end.
"#);
    assert_eq!(out, vec!["Hello Alice", "Welcome Bob"]);
}

#[test]
fn test_default_parameter_multiple() {
    let out = run_pascal(r#"
program Test;
procedure SetBox(w: Integer = 100; h: Integer = 200; color: String = 'Red');
begin
  WriteLn(w.ToString + 'x' + h.ToString + '-' + color);
end;
begin
  SetBox;
  SetBox(50);
  SetBox(50, 75, 'Blue');
end.
"#);
    assert_eq!(out, vec!["100x200-Red", "50x200-Red", "50x75-Blue"]);
}

#[test]
fn test_overload_combined_with_default_params() {
    let out = run_pascal(r#"
program Test;
procedure Process(x: Integer; msg: String = 'DefaultInt'); overload;
begin WriteLn('IntProc:' + x.ToString + '-' + msg); end;

procedure Process(x: Double; msg: String = 'DefaultFloat'); overload;
begin WriteLn('FloatProc:' + FloatToStr(x) + '-' + msg); end;

begin
  Process(10);
  Process(2.5);
end.
"#);
    assert_eq!(out, vec!["IntProc:10-DefaultInt", "FloatProc:2.5-DefaultFloat"]);
}

#[test]
fn test_default_parameter_enum_type() {
    let out = run_pascal(r#"
program Test;
type TLevel = (lvlLow, lvlMed, lvlHigh);
procedure Log(msg: String; lvl: TLevel = lvlMed);
begin
  WriteLn(msg + ':' + Ord(lvl).ToString);
end;
begin
  Log('Message1');
  Log('Message2', lvlHigh);
end.
"#);
    assert_eq!(out, vec!["Message1:1", "Message2:2"]);
}

#[test]
fn test_default_parameter_boolean_flag() {
    let out = run_pascal(r#"
program Test;
procedure Render(const text: String; uppercase: Boolean = False);
begin
  if uppercase then WriteLn(UpperCase(text))
  else WriteLn(text);
end;
begin
  Render('Pascal');
  Render('Pascal', True);
end.
"#);
    assert_eq!(out, vec!["Pascal", "PASCAL"]);
}

#[test]
fn test_overload_in_class_methods() {
    let out = run_pascal(r#"
program Test;
type TCalculator = class
  public
    function Compute(a, b: Integer): Integer; overload;
    function Compute(a, b: Double): Double; overload;
end;
function TCalculator.Compute(a, b: Integer): Integer; begin Result := a + b; end;
function TCalculator.Compute(a, b: Double): Double; begin Result := a * b; end;

var calc: TCalculator;
begin
  calc := TCalculator.Create;
  WriteLn(calc.Compute(5, 5));
  WriteLn(calc.Compute(2.5, 4.0));
  calc.Free;
end.
"#);
    assert_eq!(out, vec!["10", "10"]);
}

#[test]
fn test_overload_in_record_methods() {
    let out = run_pascal(r#"
program Test;
type TFormatter = record
  procedure Format(v: Integer); overload;
  procedure Format(const s: String); overload;
end;
procedure TFormatter.Format(v: Integer); begin WriteLn('RecInt:' + v.ToString); end;
procedure TFormatter.Format(const s: String); begin WriteLn('RecStr:' + s); end;

var f: TFormatter;
begin
  f.Format(99);
  f.Format('Text');
end.
"#);
    assert_eq!(out, vec!["RecInt:99", "RecStr:Text"]);
}

#[test]
fn test_default_parameter_constant_expression() {
    let out = run_pascal(r#"
program Test;
const DEFAULT_LIMIT = 50;
procedure LimitCheck(val: Integer; maxVal: Integer = DEFAULT_LIMIT);
begin
  WriteLn(val <= maxVal);
end;
begin
  LimitCheck(30);
  LimitCheck(70);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_overload_resolution_widening_conversion() {
    let out = run_pascal(r#"
program Test;
procedure Run(v: Int64); overload;
begin WriteLn('Int64:' + v.ToString); end;

procedure Run(v: Double); overload;
begin WriteLn('Double:' + FloatToStr(v)); end;

var b: Byte;
begin
  b := 10;
  Run(b); // Prefers Int64 widening over Double
end.
"#);
    assert_eq!(out, vec!["Int64:10"]);
}

#[test]
fn test_default_parameter_nil_pointer() {
    let out = run_pascal(r#"
program Test;
procedure Inspect(ptr: Pointer = nil);
begin
  WriteLn(ptr = nil);
end;
var x: Integer;
begin
  Inspect;
  Inspect(@x);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_overload_virtual_method_override() {
    let out = run_pascal(r#"
program Test;
type TBase = class
  public procedure Show(x: Integer); overload; virtual;
end;
type TSub = class(TBase)
  public
    procedure Show(x: Integer); overload; override;
    procedure Show(const s: String); overload;
end;
procedure TBase.Show(x: Integer); begin WriteLn('BaseInt:' + x.ToString); end;
procedure TSub.Show(x: Integer); begin WriteLn('SubInt:' + x.ToString); end;
procedure TSub.Show(const s: String); begin WriteLn('SubStr:' + s); end;

var obj: TSub;
begin
  obj := TSub.Create;
  obj.Show(10);
  obj.Show('Text');
  obj.Free;
end.
"#);
    assert_eq!(out, vec!["SubInt:10", "SubStr:Text"]);
}

#[test]
fn test_default_parameter_negative_number() {
    let out = run_pascal(r#"
program Test;
procedure Adjust(val: Integer; offset: Integer = -5);
begin
  WriteLn(val + offset);
end;
begin
  Adjust(10);
  Adjust(10, 20);
end.
"#);
    assert_eq!(out, vec!["5", "30"]);
}

#[test]
fn test_overload_procedure_pointer_resolution() {
    let out = run_pascal(r#"
program Test;
procedure Target(x: Integer); overload; begin WriteLn('ProcInt:' + x.ToString); end;
procedure Target(s: String); overload; begin WriteLn('ProcStr:' + s); end;

type TIntProc = procedure(x: Integer);
var p: TIntProc;
begin
  p := Target;
  p(100);
end.
"#);
    assert_eq!(out, vec!["ProcInt:100"]);
}

#[test]
fn test_default_parameter_empty_string() {
    let out = run_pascal(r#"
program Test;
procedure LogTag(const msg: String; const tag: String = '');
begin
  if tag = '' then WriteLn(msg)
  else WriteLn('[' + tag + '] ' + msg);
end;
begin
  LogTag('PlainMsg');
  LogTag('TaggedMsg', 'INFO');
end.
"#);
    assert_eq!(out, vec!["PlainMsg", "[INFO] TaggedMsg"]);
}

#[test]
fn test_overload_interface_methods() {
    let out = run_pascal(r#"
program Test;
type IOverloaded = interface
  ['{12341234-1234-1234-1234-123412341234}']
  procedure DoIt(x: Integer); overload;
  procedure DoIt(const s: String); overload;
end;

type TOverloadedImpl = class(TInterfacedObject, IOverloaded)
  public
    procedure DoIt(x: Integer); overload;
    procedure DoIt(const s: String); overload;
end;
procedure TOverloadedImpl.DoIt(x: Integer); begin WriteLn('IntfInt:' + x.ToString); end;
procedure TOverloadedImpl.DoIt(const s: String); begin WriteLn('IntfStr:' + s); end;

var intf: IOverloaded;
begin
  intf := TOverloadedImpl.Create;
  intf.DoIt(42);
  intf.DoIt('Data');
end.
"#);
    assert_eq!(out, vec!["IntfInt:42", "IntfStr:Data"]);
}

#[test]
fn test_default_parameter_typed_constant() {
    let out = run_pascal(r#"
program Test;
const DEFAULT_FLOAT: Double = 1.23;
procedure Scale(val: Double; factor: Double = DEFAULT_FLOAT);
begin
  WriteLn(val * factor);
end;
begin
  Scale(2.0);
end.
"#);
    assert_eq!(out, vec!["2.46"]);
}

#[test]
fn test_overload_array_parameter_variants() {
    let out = run_pascal(r#"
program Test;
procedure ProcessArray(const arr: array of Integer); overload;
begin WriteLn('IntArrayLength:' + Length(arr).ToString); end;

procedure ProcessArray(const arr: array of String); overload;
begin WriteLn('StrArrayLength:' + Length(arr).ToString); end;

begin
  ProcessArray([10, 20, 30]);
  ProcessArray(['A', 'B']);
end.
"#);
    assert_eq!(out, vec!["IntArrayLength:3", "StrArrayLength:2"]);
}

#[test]
fn test_default_parameter_char_type() {
    let out = run_pascal(r#"
program Test;
procedure Pad(const s: String; ch: Char = ' ');
begin
  WriteLn(s + ch);
end;
begin
  Pad('A');
  Pad('A', '*');
end.
"#);
    assert_eq!(out, vec!["A ", "A*"]);
}
