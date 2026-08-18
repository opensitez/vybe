use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 47: Custom Comparers & Equality Comparers (IComparer<T>, TEqualityComparer<T>)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tcomparer_default_integer_compare() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var cmp: IComparer<Integer>;
begin
  cmp := TComparer<Integer>.Default;
  WriteLn(cmp.Compare(10, 20) < 0);
  WriteLn(cmp.Compare(20, 10) > 0);
  WriteLn(cmp.Compare(15, 15) = 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE", "TRUE"]);
}

#[test]
fn test_tequalitycomparer_default_string_equals() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var eq: IEqualityComparer<String>;
begin
  eq := TEqualityComparer<String>.Default;
  WriteLn(eq.Equals('hello', 'hello'));
  WriteLn(eq.Equals('hello', 'world'));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "FALSE"]);
}

#[test]
fn test_custom_string_length_comparer() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var cmp: IComparer<String>;
begin
  cmp := TComparer<String>.Construct(
    function(const L, R: String): Integer
    begin
      Result := Length(L) - Length(R);
    end);
  WriteLn(cmp.Compare('A', 'BBB') < 0);
  WriteLn(cmp.Compare('XXXX', 'YY') > 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_custom_descending_integer_comparer() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections, Generics.Defaults;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(10); list.Add(50); list.Add(30);
  list.Sort(TComparer<Integer>.Construct(
    function(const L, R: Integer): Integer
    begin
      Result := R - L;
    end));
  WriteLn(list[0]);
  WriteLn(list[1]);
  WriteLn(list[2]);
  list.Free;
end.
"#,
    );
    assert_eq!(out, vec!["50", "30", "10"]);
}

#[test]
fn test_case_insensitive_string_equality_comparer() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils, Generics.Defaults, Generics.Collections;
var dict: TDictionary<String, Integer>;
    eq: IEqualityComparer<String>;
begin
  eq := TEqualityComparer<String>.Construct(
    function(const L, R: String): Boolean
    begin
      Result := SameText(L, R);
    end,
    function(const Value: String): Integer
    begin
      Result := HashName(LowerCase(Value));
    end);
  dict := TDictionary<String, Integer>.Create(eq);
  dict.Add('KeyName', 100);
  WriteLn(dict['keyname']);
  dict.Free;
end.
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_record_multi_field_comparer() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
type TPerson = record Name: String; Age: Integer; end;
var cmp: IComparer<TPerson>; p1, p2: TPerson;
begin
  cmp := TComparer<TPerson>.Construct(
    function(const L, R: TPerson): Integer
    begin
      Result := CompareText(L.Name, R.Name);
      if Result = 0 then Result := L.Age - R.Age;
    end);
  p1.Name := 'Alice'; p1.Age := 25;
  p2.Name := 'Alice'; p2.Age := 30;
  WriteLn(cmp.Compare(p1, p2) < 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_float_comparer_with_tolerance() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var cmp: IComparer<Real>;
begin
  cmp := TComparer<Real>.Construct(
    function(const L, R: Real): Integer
    begin
      if Abs(L - R) < 0.001 then Result := 0
      else if L < R then Result := -1
      else Result := 1;
    end);
  WriteLn(cmp.Compare(1.0001, 1.0002) = 0);
  WriteLn(cmp.Compare(1.0, 2.0) < 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_enum_custom_comparer() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
type TPriority = (pLow, pMed, pHigh);
var cmp: IComparer<TPriority>;
begin
  cmp := TComparer<TPriority>.Construct(
    function(const L, R: TPriority): Integer
    begin
      Result := Ord(L) - Ord(R);
    end);
  WriteLn(cmp.Compare(pLow, pHigh) < 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_tarray_sort_with_custom_comparer() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections, Generics.Defaults;
var arr: TArray<String>;
begin
  SetLength(arr, 3);
  arr[0] := 'banana'; arr[1] := 'apple'; arr[2] := 'cherry';
  TArray.Sort<String>(arr, TComparer<String>.Construct(
    function(const L, R: String): Integer
    begin
      Result := CompareText(L, R);
    end));
  WriteLn(arr[0]);
  WriteLn(arr[1]);
  WriteLn(arr[2]);
end.
"#,
    );
    assert_eq!(out, vec!["apple", "banana", "cherry"]);
}

#[test]
fn test_custom_class_instance_comparer() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
type TItem = class
  public Priority: Integer;
  constructor Create(P: Integer);
end;
constructor TItem.Create(P: Integer); begin Priority := P; end;

var cmp: IComparer<TItem>; i1, i2: TItem;
begin
  cmp := TComparer<TItem>.Construct(
    function(const L, R: TItem): Integer
    begin
      Result := L.Priority - R.Priority;
    end);
  i1 := TItem.Create(10); i2 := TItem.Create(20);
  WriteLn(cmp.Compare(i1, i2) < 0);
  i1.Free; i2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_gethashcode_integer_default() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var eq: IEqualityComparer<Integer>;
begin
  eq := TEqualityComparer<Integer>.Default;
  WriteLn(eq.GetHashCode(100) = eq.GetHashCode(100));
  WriteLn(eq.GetHashCode(100) <> eq.GetHashCode(200));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_gethashcode_custom_record() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
type TPoint = record X, Y: Integer; end;
var eq: IEqualityComparer<TPoint>; pt1, pt2: TPoint;
begin
  eq := TEqualityComparer<TPoint>.Construct(
    function(const L, R: TPoint): Boolean
    begin
      Result := (L.X = R.X) and (L.Y = R.Y);
    end,
    function(const Value: TPoint): Integer
    begin
      Result := Value.X xor (Value.Y shl 16);
    end);
  pt1.X := 10; pt1.Y := 20;
  pt2.X := 10; pt2.Y := 20;
  WriteLn(eq.Equals(pt1, pt2));
  WriteLn(eq.GetHashCode(pt1) = eq.GetHashCode(pt2));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_tequalitycomparer_default_boolean() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var eq: IEqualityComparer<Boolean>;
begin
  eq := TEqualityComparer<Boolean>.Default;
  WriteLn(eq.Equals(True, True));
  WriteLn(eq.Equals(True, False));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "FALSE"]);
}

#[test]
fn test_tcomparer_binarysearch() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections, Generics.Defaults;
var list: TList<Integer>; idx: Integer;
begin
  list := TList<Integer>.Create;
  list.Add(10); list.Add(20); list.Add(30);
  if list.BinarySearch(20, idx, TComparer<Integer>.Default) then
    WriteLn(idx);
  list.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_custom_comparer_reusability() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections, Generics.Defaults;
var cmp: IComparer<Integer>;
    l1, l2: TList<Integer>;
begin
  cmp := TComparer<Integer>.Construct(
    function(const L, R: Integer): Integer begin Result := L - R; end);
  l1 := TList<Integer>.Create(cmp);
  l2 := TList<Integer>.Create(cmp);
  l1.Add(2); l1.Add(1); l1.Sort;
  l2.Add(20); l2.Add(10); l2.Sort;
  WriteLn(l1[0]);
  WriteLn(l2[0]);
  l1.Free; l2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1", "10"]);
}

#[test]
fn test_custom_equality_comparer_for_boolean_flags() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var eq: IEqualityComparer<Boolean>;
begin
  eq := TEqualityComparer<Boolean>.Construct(
    function(const L, R: Boolean): Boolean begin Result := L = R; end,
    function(const Value: Boolean): Integer begin Result := Ord(Value); end);
  WriteLn(eq.Equals(False, False));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_comparer_string_case_insensitive_order() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils, Generics.Defaults;
var cmp: IComparer<String>;
begin
  cmp := TComparer<String>.Construct(
    function(const L, R: String): Integer
    begin
      Result := CompareText(L, R);
    end);
  WriteLn(cmp.Compare('apple', 'APPLE') = 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_tcomparer_real_default() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var cmp: IComparer<Real>;
begin
  cmp := TComparer<Real>.Default;
  WriteLn(cmp.Compare(1.5, 2.5) < 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_tequalitycomparer_real_default() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Defaults;
var eq: IEqualityComparer<Real>;
begin
  eq := TEqualityComparer<Real>.Default;
  WriteLn(eq.Equals(3.14, 3.14));
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_custom_comparer_even_odd_partitioning() {
    let out = run_pascal(
        r#"
program Test;
uses Generics.Collections, Generics.Defaults;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(1); list.Add(2); list.Add(3); list.Add(4);
  list.Sort(TComparer<Integer>.Construct(
    function(const L, R: Integer): Integer
    begin
      Result := (L mod 2) - (R mod 2);
    end));
  WriteLn(list[0] mod 2);
  WriteLn(list[1] mod 2);
  list.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}
