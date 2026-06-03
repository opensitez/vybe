/// Tests for nested and advanced record patterns in Pascal/Delphi:
/// Records containing other records, records with methods operating
/// on nested data, arrays of records, and record copying semantics.
use super::helpers::run_pascal;

// ===================================================================
// NESTED RECORDS
// ===================================================================

#[test]
fn nested_record_access() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TAddress = record
    Street: String;
    City: String;
  end;
  TPerson = record
    Name: String;
    Age: Integer;
    Address: TAddress;
  end;
var p: TPerson;
begin
  p.Name := 'Alice';
  p.Age := 30;
  p.Address.Street := '123 Main St';
  p.Address.City := 'Springfield';
  WriteLn(p.Name);
  WriteLn(p.Address.City);
end."#
        ),
        &["Alice", "Springfield"]
    );
}

#[test]
fn nested_record_function() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = record
    X: Integer;
    Y: Integer;
  end;
  TRect = record
    TopLeft: TPoint;
    BottomRight: TPoint;
  end;
function RectWidth(r: TRect): Integer;
begin
  Result := r.BottomRight.X - r.TopLeft.X;
end;
var r: TRect;
begin
  r.TopLeft.X := 10;
  r.TopLeft.Y := 20;
  r.BottomRight.X := 110;
  r.BottomRight.Y := 70;
  WriteLn(RectWidth(r));
end."#
        ),
        &["100"]
    );
}

#[test]
fn record_copy_semantics() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = record
    X: Integer;
    Y: Integer;
  end;
var a, b: TPoint;
begin
  a.X := 5;
  a.Y := 10;
  b := a;
  b.X := 99;
  WriteLn(a.X);
  WriteLn(b.X);
end."#
        ),
        &["5", "99"]
    );
}

// ===================================================================
// ARRAY OF RECORDS
// ===================================================================

#[test]
fn array_of_records_iterate() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TStudent = record
    Name: String;
    Score: Integer;
  end;
var students: array[1..3] of TStudent;
    i: Integer;
begin
  students[1].Name := 'Alice'; students[1].Score := 95;
  students[2].Name := 'Bob';   students[2].Score := 80;
  students[3].Name := 'Carol'; students[3].Score := 88;
  for i := 1 to 3 do
    WriteLn(students[i].Name + ': ' + IntToStr(students[i].Score));
end."#
        ),
        &["Alice: 95", "Bob: 80", "Carol: 88"]
    );
}

#[test]
fn array_of_records_max() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TItem = record
    Value: Integer;
  end;
var items: array[1..4] of TItem;
    i, maxVal: Integer;
begin
  items[1].Value := 10;
  items[2].Value := 40;
  items[3].Value := 20;
  items[4].Value := 30;
  maxVal := items[1].Value;
  for i := 2 to 4 do
    if items[i].Value > maxVal then
      maxVal := items[i].Value;
  WriteLn(maxVal);
end."#
        ),
        &["40"]
    );
}

#[test]
fn array_of_records_sum() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMeasure = record
    Value: Real;
  end;
var data: array[1..3] of TMeasure;
    total: Real;
    i: Integer;
begin
  data[1].Value := 1.5;
  data[2].Value := 2.5;
  data[3].Value := 3.0;
  total := 0;
  for i := 1 to 3 do
    total := total + data[i].Value;
  WriteLn(total);
end."#
        ),
        &["7"]
    );
}

// ===================================================================
// RECORD WITH MULTIPLE METHODS
// ===================================================================

#[test]
fn record_method_chain() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCounter = record
    Value: Integer;
    procedure Increment;
    procedure Decrement;
    function Get: Integer;
  end;
procedure TCounter.Increment;
begin
  Value := Value + 1;
end;
procedure TCounter.Decrement;
begin
  Value := Value - 1;
end;
function TCounter.Get: Integer;
begin
  Result := Value;
end;
var c: TCounter;
begin
  c.Value := 0;
  c.Increment;
  c.Increment;
  c.Increment;
  c.Decrement;
  WriteLn(c.Get);
end."#
        ),
        &["2"]
    );
}

#[test]
fn record_predicate_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TRange = record
    Lo: Integer;
    Hi: Integer;
    function Contains(v: Integer): Boolean;
  end;
function TRange.Contains(v: Integer): Boolean;
begin
  Result := (v >= Lo) and (v <= Hi);
end;
var r: TRange;
begin
  r.Lo := 10;
  r.Hi := 20;
  WriteLn(r.Contains(15));
  WriteLn(r.Contains(25));
end."#
        ),
        &["true", "false"]
    );
}

// ===================================================================
// RECORD WITH CONSTRUCTOR AND METHOD
// ===================================================================

#[test]
fn record_constructor_use() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TVector = record
    X, Y: Real;
    constructor Create(aX, aY: Real);
    function Length: Real;
  end;
constructor TVector.Create(aX, aY: Real);
begin
  X := aX;
  Y := aY;
end;
function TVector.Length: Real;
begin
  Result := Sqrt(X * X + Y * Y);
end;
var v: TVector;
begin
  v := TVector.Create(3.0, 4.0);
  WriteLn(v.Length);
end."#
        ),
        &["5"]
    );
}

#[test]
fn record_string_format_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = record
    X, Y: Integer;
    function ToString: String;
  end;
function TPoint.ToString: String;
begin
  Result := '(' + IntToStr(X) + ', ' + IntToStr(Y) + ')';
end;
var p: TPoint;
begin
  p.X := 3;
  p.Y := 7;
  WriteLn(p.ToString);
end."#
        ),
        &["(3, 7)"]
    );
}

// ===================================================================
// RECORD IN FUNCTION RESULT
// ===================================================================

#[test]
fn record_as_function_result() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMinMax = record
    Min: Integer;
    Max: Integer;
  end;
function FindMinMax(a, b, c: Integer): TMinMax;
begin
  if a < b then Result.Min := a else Result.Min := b;
  if Result.Min > c then Result.Min := c;
  if a > b then Result.Max := a else Result.Max := b;
  if Result.Max < c then Result.Max := c;
end;
var mm: TMinMax;
begin
  mm := FindMinMax(5, 2, 8);
  WriteLn(mm.Min);
  WriteLn(mm.Max);
end."#
        ),
        &["2", "8"]
    );
}

// ===================================================================
// RECORD EQUALITY COMPARISON VIA FIELDS
// ===================================================================

#[test]
fn record_field_equality() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCoord = record
    X, Y: Integer;
    function Equals(other: TCoord): Boolean;
  end;
function TCoord.Equals(other: TCoord): Boolean;
begin
  Result := (X = other.X) and (Y = other.Y);
end;
var a, b: TCoord;
begin
  a.X := 1; a.Y := 2;
  b.X := 1; b.Y := 2;
  WriteLn(a.Equals(b));
  b.Y := 3;
  WriteLn(a.Equals(b));
end."#
        ),
        &["true", "false"]
    );
}

// ===================================================================
// RECORD ACCUMULATION
// ===================================================================

#[test]
fn record_stats_accumulation() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TStats = record
    Count: Integer;
    Sum: Integer;
    function Average: Real;
  end;
function TStats.Average: Real;
begin
  if Count = 0 then Result := 0.0
  else Result := Sum / Count;
end;
var s: TStats;
    i: Integer;
begin
  s.Count := 0;
  s.Sum := 0;
  for i := 1 to 5 do
  begin
    s.Count := s.Count + 1;
    s.Sum := s.Sum + i;
  end;
  WriteLn(s.Average);
end."#
        ),
        &["3"]
    );
}

// ===================================================================
// MULTI-LEVEL RECORD PASSING
// ===================================================================

#[test]
fn record_passed_by_var() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TBox = record
    Width: Integer;
    Height: Integer;
  end;
procedure Scale(var b: TBox; factor: Integer);
begin
  b.Width := b.Width * factor;
  b.Height := b.Height * factor;
end;
var box: TBox;
begin
  box.Width := 5;
  box.Height := 3;
  Scale(box, 4);
  WriteLn(box.Width);
  WriteLn(box.Height);
end."#
        ),
        &["20", "12"]
    );
}

#[test]
fn record_swap_values() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPair = record
    First: Integer;
    Second: Integer;
    procedure Swap;
  end;
procedure TPair.Swap;
var tmp: Integer;
begin
  tmp := First;
  First := Second;
  Second := tmp;
end;
var p: TPair;
begin
  p.First := 100;
  p.Second := 200;
  p.Swap;
  WriteLn(p.First);
  WriteLn(p.Second);
end."#
        ),
        &["200", "100"]
    );
}

// ===================================================================
// RECORD HELPER PATTERN
// ===================================================================

#[test]
fn record_distance_calculation() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPoint = record
    X, Y: Real;
    function DistanceTo(other: TPoint): Real;
  end;
function TPoint.DistanceTo(other: TPoint): Real;
var dx, dy: Real;
begin
  dx := X - other.X;
  dy := Y - other.Y;
  Result := Sqrt(dx * dx + dy * dy);
end;
var p1, p2: TPoint;
begin
  p1.X := 0.0; p1.Y := 0.0;
  p2.X := 3.0; p2.Y := 4.0;
  WriteLn(p1.DistanceTo(p2));
end."#
        ),
        &["5"]
    );
}

#[test]
fn record_default_value_method() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TConfig = record
    Timeout: Integer;
    MaxRetries: Integer;
    procedure SetDefaults;
  end;
procedure TConfig.SetDefaults;
begin
  Timeout := 30;
  MaxRetries := 3;
end;
var cfg: TConfig;
begin
  cfg.SetDefaults;
  WriteLn(cfg.Timeout);
  WriteLn(cfg.MaxRetries);
end."#
        ),
        &["30", "3"]
    );
}
