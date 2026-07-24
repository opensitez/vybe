use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 84: Record Helpers & Type Extensions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_record_helper_integer_extension() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  public
    function IsEven: Boolean;
    function Square: Integer;
end;
function TIntHelper.IsEven: Boolean; begin Result := Self mod 2 = 0; end;
function TIntHelper.Square: Integer; begin Result := Self * Self; end;

var x: Integer;
begin
  x := 4;
  WriteLn(x.IsEven);
  WriteLn(x.Square);
end.
"#);
    assert_eq!(out, vec!["True", "16"]);
}

#[test]
fn test_record_helper_string_extension() {
    let out = run_pascal(r#"
program Test;
type TStrHelper = record helper for String
  public
    function Reversed: String;
end;
function TStrHelper.Reversed: String;
var i: Integer;
begin
  Result := '';
  for i := Length(Self) downto 1 do
    Result := Result + Self[i];
end;

var s: String;
begin
  s := 'Pascal';
  WriteLn(s.Reversed);
end.
"#);
    assert_eq!(out, vec!["lacsaP"]);
}

#[test]
fn test_record_helper_double_extension() {
    let out = run_pascal(r#"
program Test;
type TDoubleHelper = record helper for Double
  public
    function Half: Double;
end;
function TDoubleHelper.Half: Double; begin Result := Self / 2.0; end;

var d: Double;
begin
  d := 10.0;
  WriteLn(d.Half);
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_record_helper_boolean_extension() {
    let out = run_pascal(r#"
program Test;
type TBoolHelper = record helper for Boolean
  public
    function AsString: String;
end;
function TBoolHelper.AsString: String;
begin
  if Self then Result := 'YES' else Result := 'NO';
end;

var b: Boolean;
begin
  b := True;
  WriteLn(b.AsString);
end.
"#);
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn test_record_helper_custom_record() {
    let out = run_pascal(r#"
program Test;
type TPoint = record X, Y: Integer; end;
type TPointHelper = record helper for TPoint
  public
    function Area: Integer;
end;
function TPointHelper.Area: Integer; begin Result := Self.X * Self.Y; end;

var pt: TPoint;
begin
  pt.X := 5; pt.Y := 6;
  WriteLn(pt.Area);
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_record_helper_class_function() {
    let out = run_pascal(r#"
program Test;
type TPoint = record X, Y: Integer; end;
type TPointHelper = record helper for TPoint
  public
    class function CreatePos(AX, AY: Integer): TPoint; static;
end;
class function TPointHelper.CreatePos(AX, AY: Integer): TPoint;
begin
  Result.X := AX; Result.Y := AY;
end;

var pt: TPoint;
begin
  pt := TPoint.CreatePos(10, 20);
  WriteLn(pt.X.ToString + ',' + pt.Y.ToString);
end.
"#);
    assert_eq!(out, vec!["10,20"]);
}

#[test]
fn test_record_helper_mutating_method() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  public
    procedure DoubleSelf;
end;
procedure TIntHelper.DoubleSelf;
begin
  Self := Self * 2;
end;

var val: Integer;
begin
  val := 15;
  val.DoubleSelf;
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_record_helper_property_getter() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  private
    function GetIsPositive: Boolean;
  public
    property IsPositive: Boolean read GetIsPositive;
end;
function TIntHelper.GetIsPositive: Boolean; begin Result := Self > 0; end;

var val: Integer;
begin
  val := 42;
  WriteLn(val.IsPositive);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_record_helper_enum_extension() {
    let out = run_pascal(r#"
program Test;
type TStatus = (stPending, stActive, stDone);
type TStatusHelper = record helper for TStatus
  public
    function IsFinished: Boolean;
end;
function TStatusHelper.IsFinished: Boolean;
begin
  Result := Self = stDone;
end;

var s: TStatus;
begin
  s := stDone;
  WriteLn(s.IsFinished);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_record_helper_char_extension() {
    let out = run_pascal(r#"
program Test;
type TCharHelper = record helper for Char
  public
    function IsDigit: Boolean;
end;
function TCharHelper.IsDigit: Boolean;
begin
  Result := (Self >= '0') and (Self <= '9');
end;

var c: Char;
begin
  c := '7';
  WriteLn(c.IsDigit);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_record_helper_literal_invocation() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  public
    function IncBy(val: Integer): Integer;
end;
function TIntHelper.IncBy(val: Integer): Integer; begin Result := Self + val; end;

begin
  WriteLn(100.IncBy(25));
end.
"#);
    assert_eq!(out, vec!["125"]);
}

#[test]
fn test_record_helper_multiple_helpers_latest_wins() {
    let out = run_pascal(r#"
program Test;
type TIntHelper1 = record helper for Integer
  public function Describe: String;
end;
type TIntHelper2 = record helper for Integer
  public function Describe: String;
end;
function TIntHelper1.Describe: String; begin Result := 'Helper1:' + Self.ToString; end;
function TIntHelper2.Describe: String; begin Result := 'Helper2:' + Self.ToString; end;

var x: Integer;
begin
  x := 5;
  WriteLn(x.Describe);
end.
"#);
    assert_eq!(out, vec!["Helper2:5"]);
}

#[test]
fn test_record_helper_procedure_with_var_param() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  public
    procedure SwapWith(var other: Integer);
end;
procedure TIntHelper.SwapWith(var other: Integer);
var tmp: Integer;
begin
  tmp := Self; Self := other; other := tmp;
end;

var a, b: Integer;
begin
  a := 10; b := 20;
  a.SwapWith(b);
  WriteLn(a.ToString + ',' + b.ToString);
end.
"#);
    assert_eq!(out, vec!["20,10"]);
}

#[test]
fn test_record_helper_word_type() {
    let out = run_pascal(r#"
program Test;
type TWordHelper = record helper for Word
  public
    function HighByte: Byte;
end;
function TWordHelper.HighByte: Byte; begin Result := Hi(Self); end;

var w: Word;
begin
  w := $1234;
  WriteLn(HexStr(w.HighByte, 2));
end.
"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn test_record_helper_byte_type() {
    let out = run_pascal(r#"
program Test;
type TByteHelper = record helper for Byte
  public
    function ToHex: String;
end;
function TByteHelper.ToHex: String; begin Result := HexStr(Self, 2); end;

var b: Byte;
begin
  b := 255;
  WriteLn(b.ToHex);
end.
"#);
    assert_eq!(out, vec!["FF"]);
}

#[test]
fn test_record_helper_int64_type() {
    let out = run_pascal(r#"
program Test;
type TInt64Helper = record helper for Int64
  public
    function IsPositive: Boolean;
end;
function TInt64Helper.IsPositive: Boolean; begin Result := Self > 0; end;

var v: Int64;
begin
  v := 9000000000;
  WriteLn(v.IsPositive);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_record_helper_on_packed_record() {
    let out = run_pascal(r#"
program Test;
type TPackedData = packed record
  ID: Word;
  Value: Byte;
end;
type TPackedHelper = record helper for TPackedData
  public
    function CodeStr: String;
end;
function TPackedHelper.CodeStr: String;
begin
  Result := ID.ToString + '-' + Value.ToString;
end;

var data: TPackedData;
begin
  data.ID := 100; data.Value := 5;
  WriteLn(data.CodeStr);
end.
"#);
    assert_eq!(out, vec!["100-5"]);
}

#[test]
fn test_record_helper_chaining_calls() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  public
    function AddTen: Integer;
    function DoubleVal: Integer;
end;
function TIntHelper.AddTen: Integer; begin Result := Self + 10; end;
function TIntHelper.DoubleVal: Integer; begin Result := Self * 2; end;

var x: Integer;
begin
  x := 5;
  WriteLn(x.AddTen.DoubleVal); // (5 + 10) * 2 = 30
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_record_helper_with_overloaded_methods() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  public
    function Add(v: Integer): Integer; overload;
    function Add(v: Double): Double; overload;
end;
function TIntHelper.Add(v: Integer): Integer; begin Result := Self + v; end;
function TIntHelper.Add(v: Double): Double; begin Result := Self + v; end;

var x: Integer;
begin
  x := 10;
  WriteLn(x.Add(5));
  WriteLn(x.Add(2.5));
end.
"#);
    assert_eq!(out, vec!["15", "12.5"]);
}

#[test]
fn test_record_helper_array_element_invocation() {
    let out = run_pascal(r#"
program Test;
type TIntHelper = record helper for Integer
  public
    function Squared: Integer;
end;
function TIntHelper.Squared: Integer; begin Result := Self * Self; end;

var arr: array[0..2] of Integer;
begin
  arr[0] := 2; arr[1] := 3; arr[2] := 4;
  WriteLn(arr[1].Squared);
end.
"#);
    assert_eq!(out, vec!["9"]);
}
