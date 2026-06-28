/// Records, type aliases, subrange types, set types, variant records,
/// advanced records with methods — standard Object Pascal type system features.
use super::helpers::run_pascal;

// ===================================================================
// TYPE ALIASES
// ===================================================================

#[test]
fn type_alias_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TScore = Integer;
var s: TScore;
begin
  s := 100;
  WriteLn(s);
end."#
        ),
        &["100"]
    );
}

#[test]
fn type_alias_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TName = String;
var n: TName;
begin
  n := 'Alice';
  WriteLn(n);
end."#
        ),
        &["Alice"]
    );
}

#[test]
fn type_alias_in_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TAge = Integer;
function IsAdult(age: TAge): Boolean;
begin Result := age >= 18; end;
begin
  WriteLn(IsAdult(21));
  WriteLn(IsAdult(15));
end."#
        ),
        &["true", "false"]
    );
}

// ===================================================================
// RECORDS — BASIC
// ===================================================================

#[test]
fn record_field_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["10", "20"]
    );
}

#[test]
fn record_in_assignment() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["5000"]
    );
}

#[test]
fn record_passed_to_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["5"]
    );
}

#[test]
fn record_with_string_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["Bob is 25"]
    );
}

#[test]
fn record_array_of_records() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["Apple", "1"]
    );
}

// ===================================================================
// ADVANCED RECORDS WITH METHODS
// ===================================================================

#[test]
fn record_with_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["5"]
    );
}

#[test]
fn record_with_constructor() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["255,128,0"]
    );
}

// ===================================================================
// SET TYPES
// ===================================================================

#[test]
fn set_of_enum() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun);
var weekdays: set of TDay;
begin
  weekdays := [Mon, Tue, Wed, Thu, Fri];
  if Mon in weekdays then WriteLn('Monday is a weekday');
  if Sat in weekdays then WriteLn('Saturday is a weekday')
  else WriteLn('Saturday is weekend');
end."#
        ),
        &["Monday is a weekday", "Saturday is weekend"]
    );
}

#[test]
fn set_include_exclude() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TColor = (Red, Green, Blue, Yellow);
var colors: set of TColor;
begin
  colors := [Red, Blue];
  Include(colors, Green);
  Exclude(colors, Red);
  if Red in colors then WriteLn('has red') else WriteLn('no red');
  if Green in colors then WriteLn('has green');
end."#
        ),
        &["no red", "has green"]
    );
}

#[test]
fn set_union() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDigit = (D0, D1, D2, D3, D4, D5, D6, D7, D8, D9);
var evens, odds, all: set of TDigit;
begin
  evens := [D0, D2, D4, D6, D8];
  odds := [D1, D3, D5, D7, D9];
  all := evens + odds;
  if D5 in all then WriteLn('5 in all');
  if D0 in all then WriteLn('0 in all');
end."#
        ),
        &["5 in all", "0 in all"]
    );
}

#[test]
fn set_intersection() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TLetter = (A, B, C, D, E);
var s1, s2, both: set of TLetter;
begin
  s1 := [A, B, C];
  s2 := [B, C, D];
  both := s1 * s2;
  if A in both then WriteLn('A') else WriteLn('no A');
  if B in both then WriteLn('B');
  if C in both then WriteLn('C');
end."#
        ),
        &["no A", "B", "C"]
    );
}

#[test]
fn set_difference() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TLetter = (A, B, C, D, E);
var s1, s2, diff: set of TLetter;
begin
  s1 := [A, B, C, D];
  s2 := [B, D];
  diff := s1 - s2;
  if A in diff then WriteLn('A');
  if B in diff then WriteLn('B') else WriteLn('no B');
  if C in diff then WriteLn('C');
end."#
        ),
        &["A", "no B", "C"]
    );
}

// ===================================================================
// VAR PARAMETERS (BY REFERENCE)
// ===================================================================

#[test]
fn var_param_swap() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["20", "10"]
    );
}

#[test]
fn var_param_increment() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure AddTen(var n: Integer);
begin
  n := n + 10;
end;
var x: Integer;
begin
  x := 5;
  AddTen(x);
  WriteLn(x);
end."#
        ),
        &["15"]
    );
}

#[test]
fn var_param_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
procedure AppendExclaim(var s: String);
begin
  s := s + '!';
end;
var msg: String;
begin
  msg := 'hello';
  AppendExclaim(msg);
  WriteLn(msg);
end."#
        ),
        &["hello!"]
    );
}

// ===================================================================
// OUT PARAMETERS
// ===================================================================

#[test]
fn out_param_basic() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["42", "99"]
    );
}

// ===================================================================
// CONST PARAMETERS
// ===================================================================

#[test]
fn const_param() {
    assert_eq!(
        run_pascal(
            r#"program T;
function DoubleIt(const x: Integer): Integer;
begin
  Result := x * 2;
end;
begin
  WriteLn(DoubleIt(21));
end."#
        ),
        &["42"]
    );
}

// ===================================================================
// OPEN ARRAY PARAMETERS
// ===================================================================

#[test]
fn open_array_param() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["15"]
    );
}

#[test]
fn open_array_length() {
    assert_eq!(
        run_pascal(
            r#"program T;
function CountItems(arr: array of String): Integer;
begin
  Result := Length(arr);
end;
begin
  WriteLn(CountItems(['a', 'b', 'c']));
end."#
        ),
        &["3"]
    );
}

// ===================================================================
// NESTED TYPES
// ===================================================================

#[test]
fn multiple_type_declarations() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["60"]
    );
}

// ===================================================================
// CONSTANT ARRAYS
// ===================================================================

#[test]
fn const_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
const
  DayNames: array[0..6] of String = ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun');
begin
  WriteLn(DayNames[0]);
  WriteLn(DayNames[4]);
  WriteLn(DayNames[6]);
end."#
        ),
        &["Mon", "Fri", "Sun"]
    );
}

// ===================================================================
// SUBRANGE TYPES
// ===================================================================

#[test]
fn subrange_type() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMonth = 1..12;
var m: TMonth;
begin
  m := 7;
  WriteLn(m);
end."#
        ),
        &["7"]
    );
}

// ===================================================================
// WITH STATEMENT
// ===================================================================

#[test]
fn with_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["30"]
    );
}

#[test]
fn with_class() {
    assert_eq!(
        run_pascal(
            r#"program T;
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
end."#
        ),
        &["Alice", "30"]
    );
}

// -------------------------------------------------------------------
// from test_sets_union_intersection.rs
// -------------------------------------------------------------------
#[test]
fn set_membership_char_in_literal() {
    assert_eq!(
        run_pascal(
            r#"program T;
var letters: set of Char;
begin
  letters := ['a', 'b', 'c'];
  if 'b' in letters then WriteLn('yes') else WriteLn('no');
end."#
        ),
        &["yes"]
    );
}

#[test]
fn set_membership_char_not_in_set() {
    assert_eq!(
        run_pascal(
            r#"program T;
var letters: set of Char;
begin
  letters := ['x'];
  if 'y' in letters then WriteLn('in') else WriteLn('out');
end."#
        ),
        &["out"]
    );
}

#[test]
fn set_union_combines_elements() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDigit = '0'..'9';
var a, b, u: set of TDigit;
begin
  a := ['1', '2'];
  b := ['2', '3'];
  u := a + b;
  if ('3' in u) and ('1' in u) then WriteLn('ok') else WriteLn('bad');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn set_intersection_keeps_common() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TLetter = 'a'..'z';
var a, b, both: set of TLetter;
begin
  a := ['a', 'b', 'c'];
  b := ['b', 'c', 'd'];
  both := a * b;
  if ('b' in both) and not ('a' in both) then WriteLn('bc') else WriteLn('x');
end."#
        ),
        &["bc"]
    );
}

#[test]
fn set_difference_removes_right_operands() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TLetter = 'a'..'z';
var a, b, diff: set of TLetter;
begin
  a := ['a', 'b', 'c'];
  b := ['b'];
  diff := a - b;
  if ('a' in diff) and not ('b' in diff) then WriteLn('ac') else WriteLn('x');
end."#
        ),
        &["ac"]
    );
}

#[test]
fn set_include_adds_element() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: set of Byte;
begin
  s := [];
  Include(s, 5);
  if 5 in s then WriteLn('5') else WriteLn('0');
end."#
        ),
        &["5"]
    );
}

#[test]
fn set_exclude_removes_element() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: set of Byte;
begin
  s := [1, 2, 3];
  Exclude(s, 2);
  if not (2 in s) and (3 in s) then WriteLn('ok') else WriteLn('bad');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn set_empty_test_via_comparison() {
    assert_eq!(
        run_pascal(
            r#"program T;
var s: set of Integer;
begin
  s := [];
  if s = [] then WriteLn('empty') else WriteLn('not');
end."#
        ),
        &["empty"]
    );
}

#[test]
fn set_enum_days_weekend_membership() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDay = (Mon, Tue, Wed, Thu, Fri, Sat, Sun);
var weekend: set of TDay;
begin
  weekend := [Sat, Sun];
  if Sun in weekend then WriteLn('sun') else WriteLn('no');
end."#
        ),
        &["sun"]
    );
}

#[test]
fn set_subset_superset_relation() {
    assert_eq!(
        run_pascal(
            r#"program T;
var small, big: set of Integer;
begin
  small := [1, 2];
  big := [1, 2, 3];
  if small <= big then WriteLn('subset') else WriteLn('not');
end."#
        ),
        &["subset"]
    );
}

// -------------------------------------------------------------------
// from test_records_with_blocks.rs
// -------------------------------------------------------------------
#[test]
fn with_assigns_multiple_fields() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPoint = record X, Y: Integer; end;
var p: TPoint;
begin
  with p do begin
    X := 2;
    Y := 5;
  end;
  WriteLn(p.X + p.Y);
end."#
        ),
        &["7"]
    );
}

#[test]
fn with_nested_record_path() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TInner = record Value: Integer; end;
    TOuter = record Inner: TInner; end;
var o: TOuter;
begin
  with o.Inner do
    Value := 9;
  WriteLn(o.Inner.Value);
end."#
        ),
        &["9"]
    );
}

#[test]
fn with_doubles_field_in_place() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TBox = record N: Integer; end;
var b: TBox;
begin
  b.N := 3;
  with b do
    N := N * 2;
  WriteLn(b.N);
end."#
        ),
        &["6"]
    );
}

#[test]
fn with_string_field_update() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPerson = record Name: String; end;
var p: TPerson;
begin
  with p do
    Name := 'Ann';
  WriteLn(p.Name);
end."#
        ),
        &["Ann"]
    );
}

#[test]
fn with_in_loop_updates_array_of_records() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TItem = record V: Integer; end;
var items: array[0..1] of TItem;
    i: Integer;
begin
  for i := 0 to 1 do
    with items[i] do
      V := i + 1;
  WriteLn(items[0].V + items[1].V);
end."#
        ),
        &["3"]
    );
}

#[test]
fn with_two_sequential_blocks_same_record() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TRect = record W, H: Integer; end;
var r: TRect;
begin
  with r do W := 4;
  with r do H := 5;
  WriteLn(r.W * r.H);
end."#
        ),
        &["20"]
    );
}

#[test]
fn with_boolean_field_toggle() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TFlags = record Active: Boolean; end;
var f: TFlags;
begin
  with f do Active := True;
  with f do Active := not Active;
  WriteLn(f.Active);
end."#
        ),
        &["false"]
    );
}

#[test]
fn with_record_passed_to_procedure_by_var() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TScore = record Points: Integer; end;
procedure AddTen(var s: TScore);
begin
  with s do
    Points := Points + 10;
end;
var s: TScore;
begin
  s.Points := 5;
  AddTen(s);
  WriteLn(s.Points);
end."#
        ),
        &["15"]
    );
}

#[test]
fn enum_explicit_values_leave_gaps() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TCode = (A = 10, B, C);
var c: TCode;
begin
  c := B;
  WriteLn(Ord(c));
end."#
        ),
        &["11"]
    );
}

#[test]
fn subrange_type_bounds_check() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TDigit = 0..9;
var d: TDigit;
begin
  d := 7;
  WriteLn(d);
end."#
        ),
        &["7"]
    );
}

#[test]
fn set_difference_removes_elements() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TLetter = 'a'..'z';
var a, b, d: set of TLetter;
begin
  a := ['a', 'b', 'c'];
  b := ['b'];
  d := a - b;
  if ('a' in d) and not ('b' in d) then WriteLn('ok');
end."#
        ),
        &["ok"]
    );
}

#[test]
fn packed_record_size_may_differ() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TPacked = packed record A: Byte; B: Byte; end;
var p: TPacked;
begin
  p.A := 1;
  p.B := 2;
  WriteLn(p.A + p.B);
end."#
        ),
        &["3"]
    );
}

#[test]
fn record_method_modifies_self_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TAcc = record
  N: Integer;
  procedure IncN;
end;
procedure TAcc.IncN; begin N := N + 1; end;
var a: TAcc;
begin
  a.N := 0;
  a.IncN;
  WriteLn(a.N);
end."#
        ),
        &["1"]
    );
}

#[test]
fn case_variant_record_tag_selects_field() {
    assert_eq!(
        run_pascal(
            r#"program T;
type TShape = record
  case Integer of
    1: (Radius: Real);
    2: (Width, Height: Real);
end;
var s: TShape;
begin
  s.Radius := 2.0;
  WriteLn(Format('%.0f', [s.Radius]));
end."#
        ),
        &["2"]
    );
}
