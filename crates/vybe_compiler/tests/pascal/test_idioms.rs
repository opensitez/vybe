/// Common Object Pascal idioms and patterns that Delphi/FPC developers
/// use daily: properties, with statement, nested functions, Write vs WriteLn,
/// Assigned checks, nil handling, chained field access, common algorithms.

use super::helpers::run_pascal;

// ===================================================================
// PROPERTIES — READ/WRITE
// ===================================================================

#[test] fn property_read_write_basic() {
    assert_eq!(run_pascal(r#"program T;
type
  TPerson = class
  private
    FName: String;
  public
    constructor Create(name: String);
    property Name: String read FName write FName;
  end;

constructor TPerson.Create(name: String);
begin FName := name; end;

var p: TPerson;
begin
  p := TPerson.Create('Alice');
  WriteLn(p.Name);
  p.Name := 'Bob';
  WriteLn(p.Name);
end."#), &["Alice", "Bob"]);
}

#[test] fn property_with_getter_setter() {
    assert_eq!(run_pascal(r#"program T;
type
  TTemperature = class
  private
    FCelsius: Real;
    function GetFahrenheit: Real;
    procedure SetFahrenheit(val: Real);
  public
    constructor Create;
    property Celsius: Real read FCelsius write FCelsius;
    property Fahrenheit: Real read GetFahrenheit write SetFahrenheit;
  end;

constructor TTemperature.Create; begin FCelsius := 0; end;
function TTemperature.GetFahrenheit: Real;
begin Result := FCelsius * 9 / 5 + 32; end;
procedure TTemperature.SetFahrenheit(val: Real);
begin FCelsius := (val - 32) * 5 / 9; end;

var t: TTemperature;
begin
  t := TTemperature.Create;
  t.Celsius := 100;
  WriteLn(t.Fahrenheit);
end."#), &["212"]);
}

#[test] fn property_readonly() {
    assert_eq!(run_pascal(r#"program T;
type
  TCounter = class
  private
    FCount: Integer;
  public
    constructor Create;
    procedure Increment;
    property Count: Integer read FCount;
  end;

constructor TCounter.Create; begin FCount := 0; end;
procedure TCounter.Increment; begin FCount := FCount + 1; end;

var c: TCounter;
begin
  c := TCounter.Create;
  c.Increment;
  c.Increment;
  c.Increment;
  WriteLn(c.Count);
end."#), &["3"]);
}

// ===================================================================
// ASSIGNED CHECKS
// ===================================================================

#[test] fn assigned_check_on_object() {
    assert_eq!(run_pascal(r#"program T;
type TFoo = class
  public constructor Create;
end;
constructor TFoo.Create; begin end;

var f: TFoo;
begin
  f := nil;
  if not Assigned(f) then WriteLn('nil');
  f := TFoo.Create;
  if Assigned(f) then WriteLn('assigned');
end."#), &["nil", "assigned"]);
}

#[test] fn assigned_check_before_method_call() {
    assert_eq!(run_pascal(r#"program T;
type TLogger = class
  public
    constructor Create;
    procedure Log(msg: String);
end;
constructor TLogger.Create; begin end;
procedure TLogger.Log(msg: String); begin WriteLn(msg); end;

var logger: TLogger;
begin
  logger := nil;
  if Assigned(logger) then logger.Log('should not print');
  logger := TLogger.Create;
  if Assigned(logger) then logger.Log('hello');
end."#), &["hello"]);
}

// ===================================================================
// CHAINED FIELD ACCESS
// ===================================================================

#[test] fn chained_field_access() {
    assert_eq!(run_pascal(r#"program T;
type
  TAddress = class
  public
    FCity: String;
    constructor Create(city: String);
  end;
  TPerson = class
  public
    FName: String;
    FAddress: TAddress;
    constructor Create(name, city: String);
  end;

constructor TAddress.Create(city: String); begin FCity := city; end;
constructor TPerson.Create(name, city: String);
begin FName := name; FAddress := TAddress.Create(city); end;

var p: TPerson;
begin
  p := TPerson.Create('Alice', 'Paris');
  WriteLn(p.FName + ' lives in ' + p.FAddress.FCity);
end."#), &["Alice lives in Paris"]);
}

#[test] fn deep_chained_access() {
    assert_eq!(run_pascal(r#"program T;
type
  TC = class
  public FVal: Integer; constructor Create(v: Integer);
  end;
  TB = class
  public FC: TC; constructor Create(v: Integer);
  end;
  TA = class
  public FB: TB; constructor Create(v: Integer);
  end;

constructor TC.Create(v: Integer); begin FVal := v; end;
constructor TB.Create(v: Integer); begin FC := TC.Create(v); end;
constructor TA.Create(v: Integer); begin FB := TB.Create(v); end;

var a: TA;
begin
  a := TA.Create(42);
  WriteLn(a.FB.FC.FVal);
end."#), &["42"]);
}

// ===================================================================
// COMMON PASCAL PATTERNS
// ===================================================================

#[test] fn swap_idiom() {
    assert_eq!(run_pascal(r#"program T;
procedure Swap(var a, b: Integer);
var t: Integer;
begin t := a; a := b; b := t; end;

var x, y: Integer;
begin
  x := 1; y := 2;
  Swap(x, y);
  WriteLn(x); WriteLn(y);
end."#), &["2", "1"]);
}

#[test] fn accumulator_pattern() {
    assert_eq!(run_pascal(r#"program T;
function Sum(arr: array of Integer): Integer;
var n: Integer;
begin
  Result := 0;
  for n in arr do Result := Result + n;
end;
begin
  WriteLn(Sum([1, 2, 3, 4, 5]));
end."#), &["15"]);
}

#[test] fn guard_clause_pattern() {
    assert_eq!(run_pascal(r#"program T;
function Describe(n: Integer): String;
begin
  if n < 0 then begin Result := 'negative'; Exit; end;
  if n = 0 then begin Result := 'zero'; Exit; end;
  Result := 'positive';
end;
begin
  WriteLn(Describe(-5));
  WriteLn(Describe(0));
  WriteLn(Describe(10));
end."#), &["negative", "zero", "positive"]);
}

#[test] fn builder_pattern() {
    assert_eq!(run_pascal(r#"program T;
type
  TBuilder = class
  public
    FParts: String;
    constructor Create;
    function AddPart(part: String): TBuilder;
    function Build: String;
  end;

constructor TBuilder.Create; begin FParts := ''; end;
function TBuilder.AddPart(part: String): TBuilder;
begin
  if Length(FParts) > 0 then FParts := FParts + ', ';
  FParts := FParts + part;
  Result := Self;
end;
function TBuilder.Build: String; begin Result := '[' + FParts + ']'; end;

var b: TBuilder;
begin
  b := TBuilder.Create;
  WriteLn(b.AddPart('A').AddPart('B').AddPart('C').Build());
end."#), &["[A, B, C]"]);
}

// ===================================================================
// WRITE VS WRITELN
// ===================================================================

#[test] fn write_no_newline() {
    assert_eq!(run_pascal(r#"program T;
begin
  Write('Hello ');
  Write('World');
  WriteLn;
end."#), &["Hello World"]);
}

#[test] fn write_then_writeln() {
    assert_eq!(run_pascal(r#"program T;
var i: Integer;
begin
  for i := 1 to 5 do Write(IntToStr(i) + ' ');
  WriteLn;
end."#), &["1 2 3 4 5 "]);
}

// ===================================================================
// NESTED FUNCTION PATTERNS
// ===================================================================

#[test] fn nested_helper_function() {
    assert_eq!(run_pascal(r#"program T;
function ProcessList(arr: array of Integer): String;
  function FormatItem(n: Integer): String;
  begin
    if n mod 2 = 0 then Result := IntToStr(n) + '(even)'
    else Result := IntToStr(n) + '(odd)';
  end;
var i: Integer;
begin
  Result := '';
  for i := 0 to High(arr) do
  begin
    if i > 0 then Result := Result + ', ';
    Result := Result + FormatItem(arr[i]);
  end;
end;
begin
  WriteLn(ProcessList([1, 2, 3, 4]));
end."#), &["1(odd), 2(even), 3(odd), 4(even)"]);
}

// ===================================================================
// DEFAULT INITIALIZED VARIABLES
// ===================================================================

#[test] fn default_init_local_vars() {
    assert_eq!(run_pascal(r#"program T;
var i: Integer;
var s: String;
var b: Boolean;
var r: Real;
begin
  WriteLn(i);
  WriteLn(Length(s));
  WriteLn(b);
  WriteLn(r);
end."#), &["0", "0", "false", "0"]);
}

// ===================================================================
// MULTI-RETURN VIA VAR PARAMS
// ===================================================================

#[test] fn multi_return_via_var() {
    assert_eq!(run_pascal(r#"program T;
procedure DivMod(a, b: Integer; var quotient, remainder: Integer);
begin
  quotient := a div b;
  remainder := a mod b;
end;

var q, r: Integer;
begin
  DivMod(17, 5, q, r);
  WriteLn(q);
  WriteLn(r);
end."#), &["3", "2"]);
}

// ===================================================================
// BOOLEAN EXPRESSION PATTERNS
// ===================================================================

#[test] fn boolean_function_in_condition() {
    assert_eq!(run_pascal(r#"program T;
function InRange(val, lo, hi: Integer): Boolean;
begin
  Result := (val >= lo) and (val <= hi);
end;
begin
  if InRange(5, 1, 10) then WriteLn('in range');
  if not InRange(15, 1, 10) then WriteLn('out of range');
end."#), &["in range", "out of range"]);
}

// ===================================================================
// CASE STATEMENT PATTERNS
// ===================================================================

#[test] fn case_with_string() {
    assert_eq!(run_pascal(r#"program T;
var cmd: String;
begin
  cmd := 'hello';
  case cmd of
    'hello': WriteLn('greeting');
    'bye': WriteLn('farewell');
  else
    WriteLn('unknown');
  end;
end."#), &["greeting"]);
}

#[test] fn case_with_begin_end_blocks() {
    assert_eq!(run_pascal(r#"program T;
var x: Integer;
begin
  x := 2;
  case x of
    1: begin
      WriteLn('one');
      WriteLn('uno');
    end;
    2: begin
      WriteLn('two');
      WriteLn('dos');
    end;
    3: begin
      WriteLn('three');
    end;
  end;
end."#), &["two", "dos"]);
}

// ===================================================================
// ITERATOR / VISITOR PATTERN
// ===================================================================

#[test] fn visitor_pattern() {
    assert_eq!(run_pascal(r#"program T;
type
  TNode = class
  public
    FValue: Integer;
    constructor Create(v: Integer);
  end;

constructor TNode.Create(v: Integer); begin FValue := v; end;

procedure Visit(nodes: array of TNode; action: procedure(n: TNode));
var i: Integer;
begin
  for i := 0 to High(nodes) do action(nodes[i]);
end;

begin
  Visit(
    [TNode.Create(1), TNode.Create(2), TNode.Create(3)],
    procedure(n: TNode) begin WriteLn(n.FValue * 10); end
  );
end."#), &["10", "20", "30"]);
}
