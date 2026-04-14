/// Records, type aliases, subrange types, set types, variant records,
/// advanced records with methods — standard Object Pascal type system features.

use super::helpers::run_pascal;

// ===================================================================
// TYPE ALIASES
// ===================================================================

#[test] fn type_alias_integer() {
    assert_eq!(run_pascal(r#"program T;
type TScore = Integer;
var s: TScore;
begin
  s := 100;
  WriteLn(s);
end."#), &["100"]);
}

#[test] fn type_alias_string() {
    assert_eq!(run_pascal(r#"program T;
type TName = String;
var n: TName;
begin
  n := 'Alice';
  WriteLn(n);
end."#), &["Alice"]);
}

#[test] fn type_alias_in_function() {
    assert_eq!(run_pascal(r#"program T;
type TAge = Integer;
function IsAdult(age: TAge): Boolean;
begin Result := age >= 18; end;
begin
  WriteLn(IsAdult(21));
  WriteLn(IsAdult(15));
end."#), &["true", "false"]);
}

// ===================================================================
// RECORDS — BASIC
// ===================================================================

#[test] fn record_field_access() {
    assert_eq!(run_pascal(r#"program T;
type
  TPoint = record
    X: Integer;
    Y: Integer;
  end;
var p: TPoint;
begin
  p.X := 10;
  p.Y := 20;
  WriteLn(p.X);
  WriteLn(p.Y);
end."#), &["10", "20"]);
}

#[test] fn record_in_assignment() {
    assert_eq!(run_pascal(r#"program T;
type
  TRect = record
    Left: Integer;
    Top: Integer;
    Width: Integer;
    Height: Integer;
  end;
var r: TRect;
begin
  r.Left := 0;
  r.Top := 0;
  r.Width := 100;
  r.Height := 50;
  WriteLn(r.Width * r.Height);
end."#), &["5000"]);
}

#[test] fn record_passed_to_function() {
    assert_eq!(run_pascal(r#"program T;
type
  TPoint = record
    X: Integer;
    Y: Integer;
  end;
function DistFromOrigin(p: TPoint): Real;
begin
  Result := Sqrt(p.X * p.X + p.Y * p.Y);
end;
var pt: TPoint;
begin
  pt.X := 3;
  pt.Y := 4;
  WriteLn(DistFromOrigin(pt));
end."#), &["5"]);
}

#[test] fn record_with_string_field() {
    assert_eq!(run_pascal(r#"program T;
type
  TPerson = record
    Name: String;
    Age: Integer;
  end;
var p: TPerson;
begin
  p.Name := 'Bob';
  p.Age := 25;
  WriteLn(p.Name + ' is ' + IntToStr(p.Age));
end."#), &["Bob is 25"]);
}

#[test] fn record_array_of_records() {
    assert_eq!(run_pascal(r#"program T;
type
  TItem = record
    Name: String;
    Price: Integer;
  end;
var items: array of TItem;
var item: TItem;
begin
  item.Name := 'Apple';
  item.Price := 1;
  items := [item];
  WriteLn(items[0].Name);
  WriteLn(items[0].Price);
end."#), &["Apple", "1"]);
}

// ===================================================================
// ADVANCED RECORDS WITH METHODS
// ===================================================================

#[test] fn record_with_method() {
    assert_eq!(run_pascal(r#"program T;
type
  TVector = record
    X: Real;
    Y: Real;
    function Length: Real;
  end;

function TVector.Length: Real;
begin
  Result := Sqrt(X * X + Y * Y);
end;

var v: TVector;
begin
  v.X := 3.0;
  v.Y := 4.0;
  WriteLn(v.Length());
end."#), &["5"]);
}

#[test] fn record_with_constructor() {
    assert_eq!(run_pascal(r#"program T;
type
  TColor = record
    R: Integer;
    G: Integer;
    B: Integer;
    constructor Create(aR, aG, aB: Integer);
    function ToString: String;
  end;

constructor TColor.Create(aR, aG, aB: Integer);
begin R := aR; G := aG; B := aB; end;

function TColor.ToString: String;
begin Result := IntToStr(R) + ',' + IntToStr(G) + ',' + IntToStr(B); end;

var c: TColor;
begin
  c := TColor.Create(255, 128, 0);
  WriteLn(c.ToString());
end."#), &["255,128,0"]);
}

// ===================================================================
// SET TYPES
// ===================================================================

#[test] fn set_of_enum() {
    assert_eq!(run_pascal(r#"program T;
type
  TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun);
var weekdays: set of TDay;
begin
  weekdays := [Mon, Tue, Wed, Thu, Fri];
  if Mon in weekdays then WriteLn('Monday is a weekday');
  if Sat in weekdays then WriteLn('Saturday is a weekday')
  else WriteLn('Saturday is weekend');
end."#), &["Monday is a weekday", "Saturday is weekend"]);
}

#[test] fn set_include_exclude() {
    assert_eq!(run_pascal(r#"program T;
type TColor = (Red, Green, Blue, Yellow);
var colors: set of TColor;
begin
  colors := [Red, Blue];
  Include(colors, Green);
  Exclude(colors, Red);
  if Red in colors then WriteLn('has red') else WriteLn('no red');
  if Green in colors then WriteLn('has green');
end."#), &["no red", "has green"]);
}

#[test] fn set_union() {
    assert_eq!(run_pascal(r#"program T;
type TDigit = (D0, D1, D2, D3, D4, D5, D6, D7, D8, D9);
var evens, odds, all: set of TDigit;
begin
  evens := [D0, D2, D4, D6, D8];
  odds := [D1, D3, D5, D7, D9];
  all := evens + odds;
  if D5 in all then WriteLn('5 in all');
  if D0 in all then WriteLn('0 in all');
end."#), &["5 in all", "0 in all"]);
}

#[test] fn set_intersection() {
    assert_eq!(run_pascal(r#"program T;
type TLetter = (A, B, C, D, E);
var s1, s2, both: set of TLetter;
begin
  s1 := [A, B, C];
  s2 := [B, C, D];
  both := s1 * s2;
  if A in both then WriteLn('A') else WriteLn('no A');
  if B in both then WriteLn('B');
  if C in both then WriteLn('C');
end."#), &["no A", "B", "C"]);
}

#[test] fn set_difference() {
    assert_eq!(run_pascal(r#"program T;
type TLetter = (A, B, C, D, E);
var s1, s2, diff: set of TLetter;
begin
  s1 := [A, B, C, D];
  s2 := [B, D];
  diff := s1 - s2;
  if A in diff then WriteLn('A');
  if B in diff then WriteLn('B') else WriteLn('no B');
  if C in diff then WriteLn('C');
end."#), &["A", "no B", "C"]);
}

// ===================================================================
// VAR PARAMETERS (BY REFERENCE)
// ===================================================================

#[test] fn var_param_swap() {
    assert_eq!(run_pascal(r#"program T;
procedure Swap(var a, b: Integer);
var tmp: Integer;
begin
  tmp := a;
  a := b;
  b := tmp;
end;
var x, y: Integer;
begin
  x := 10; y := 20;
  Swap(x, y);
  WriteLn(x);
  WriteLn(y);
end."#), &["20", "10"]);
}

#[test] fn var_param_increment() {
    assert_eq!(run_pascal(r#"program T;
procedure AddTen(var n: Integer);
begin
  n := n + 10;
end;
var x: Integer;
begin
  x := 5;
  AddTen(x);
  WriteLn(x);
end."#), &["15"]);
}

#[test] fn var_param_string() {
    assert_eq!(run_pascal(r#"program T;
procedure AppendExclaim(var s: String);
begin
  s := s + '!';
end;
var msg: String;
begin
  msg := 'hello';
  AppendExclaim(msg);
  WriteLn(msg);
end."#), &["hello!"]);
}

// ===================================================================
// OUT PARAMETERS
// ===================================================================

#[test] fn out_param_basic() {
    assert_eq!(run_pascal(r#"program T;
procedure GetValues(out a, b: Integer);
begin
  a := 42;
  b := 99;
end;
var x, y: Integer;
begin
  GetValues(x, y);
  WriteLn(x);
  WriteLn(y);
end."#), &["42", "99"]);
}

// ===================================================================
// CONST PARAMETERS
// ===================================================================

#[test] fn const_param() {
    assert_eq!(run_pascal(r#"program T;
function DoubleIt(const x: Integer): Integer;
begin
  Result := x * 2;
end;
begin
  WriteLn(DoubleIt(21));
end."#), &["42"]);
}

// ===================================================================
// OPEN ARRAY PARAMETERS
// ===================================================================

#[test] fn open_array_param() {
    assert_eq!(run_pascal(r#"program T;
function SumAll(arr: array of Integer): Integer;
var i, s: Integer;
begin
  s := 0;
  for i := Low(arr) to High(arr) do
    s := s + arr[i];
  Result := s;
end;
begin
  WriteLn(SumAll([1, 2, 3, 4, 5]));
end."#), &["15"]);
}

#[test] fn open_array_length() {
    assert_eq!(run_pascal(r#"program T;
function CountItems(arr: array of String): Integer;
begin
  Result := Length(arr);
end;
begin
  WriteLn(CountItems(['a', 'b', 'c']));
end."#), &["3"]);
}

// ===================================================================
// NESTED TYPES
// ===================================================================

#[test] fn multiple_type_declarations() {
    assert_eq!(run_pascal(r#"program T;
type
  TDirection = (North, South, East, West);
  TSpeed = Integer;
  TMovement = record
    Dir: TDirection;
    Speed: TSpeed;
  end;
var m: TMovement;
begin
  m.Dir := East;
  m.Speed := 60;
  WriteLn(m.Speed);
end."#), &["60"]);
}

// ===================================================================
// CONSTANT ARRAYS
// ===================================================================

#[test] fn const_array() {
    assert_eq!(run_pascal(r#"program T;
const
  DayNames: array[0..6] of String = ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun');
begin
  WriteLn(DayNames[0]);
  WriteLn(DayNames[4]);
  WriteLn(DayNames[6]);
end."#), &["Mon", "Fri", "Sun"]);
}

// ===================================================================
// SUBRANGE TYPES
// ===================================================================

#[test] fn subrange_type() {
    assert_eq!(run_pascal(r#"program T;
type
  TMonth = 1..12;
var m: TMonth;
begin
  m := 7;
  WriteLn(m);
end."#), &["7"]);
}

// ===================================================================
// WITH STATEMENT
// ===================================================================

#[test] fn with_record() {
    assert_eq!(run_pascal(r#"program T;
type
  TPoint = record
    X: Integer;
    Y: Integer;
  end;
var p: TPoint;
begin
  with p do
  begin
    X := 10;
    Y := 20;
  end;
  WriteLn(p.X + p.Y);
end."#), &["30"]);
}

#[test] fn with_class() {
    assert_eq!(run_pascal(r#"program T;
type
  TPerson = class
  public
    FName: String;
    FAge: Integer;
    constructor Create;
  end;

constructor TPerson.Create; begin end;

var p: TPerson;
begin
  p := TPerson.Create;
  with p do
  begin
    FName := 'Alice';
    FAge := 30;
  end;
  WriteLn(p.FName);
  WriteLn(p.FAge);
end."#), &["Alice", "30"]);
}
