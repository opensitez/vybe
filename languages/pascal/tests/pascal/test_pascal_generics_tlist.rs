use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 41: Generic Containers (TList<T>)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tlist_integer_add_and_count() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(10);
  list.Add(20);
  list.Add(30);
  WriteLn(list.Count);
  WriteLn(list[0]);
  WriteLn(list[2]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["3", "10", "30"]);
}

#[test]
fn test_tlist_string_add_and_insert() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<String>;
begin
  list := TList<String>.Create;
  list.Add('First');
  list.Add('Third');
  list.Insert(1, 'Second');
  WriteLn(list[0]);
  WriteLn(list[1]);
  WriteLn(list[2]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["First", "Second", "Third"]);
}

#[test]
fn test_tlist_delete_and_remove() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(100); list.Add(200); list.Add(300);
  list.Delete(1);
  WriteLn(list.Count);
  WriteLn(list[1]);
  list.Remove(300);
  WriteLn(list.Count);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["2", "300", "1"]);
}

#[test]
fn test_tlist_contains_and_indexof() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<String>;
begin
  list := TList<String>.Create;
  list.Add('Apple'); list.Add('Banana'); list.Add('Cherry');
  WriteLn(list.Contains('Banana'));
  WriteLn(list.IndexOf('Cherry'));
  WriteLn(list.IndexOf('Orange'));
  list.Free;
end.
"#);
    assert_eq!(out, vec!["True", "2", "-1"]);
}

#[test]
fn test_tlist_clear_resets_list() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(1); list.Add(2);
  list.Clear;
  WriteLn(list.Count);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tlist_sort_default() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(50); list.Add(10); list.Add(30);
  list.Sort;
  WriteLn(list[0]);
  WriteLn(list[1]);
  WriteLn(list[2]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["10", "30", "50"]);
}

#[test]
fn test_tlist_toarray_conversion() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
    arr: TArray<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(5); list.Add(15);
  arr := list.ToArray;
  WriteLn(Length(arr));
  WriteLn(arr[0]);
  WriteLn(arr[1]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["2", "5", "15"]);
}

#[test]
fn test_tlist_for_in_loop_iteration() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<String>;
    item: String;
begin
  list := TList<String>.Create;
  list.Add('A'); list.Add('B'); list.Add('C');
  for item in list do
    WriteLn(item);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["A", "B", "C"]);
}

#[test]
fn test_tlist_record_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TPoint = record X, Y: Integer; end;
var list: TList<TPoint>; pt: TPoint;
begin
  list := TList<TPoint>.Create;
  pt.X := 10; pt.Y := 20;
  list.Add(pt);
  WriteLn(list[0].X);
  WriteLn(list[0].Y);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_tlist_capacity_preallocation() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Capacity := 100;
  WriteLn(list.Capacity >= 100);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tlist_trimexcess() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Capacity := 100;
  list.Add(1); list.Add(2);
  list.TrimExcess;
  WriteLn(list.Capacity);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_tlist_reverse() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(1); list.Add(2); list.Add(3);
  list.Reverse;
  WriteLn(list[0]);
  WriteLn(list[2]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn test_tlist_exchange_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<String>;
begin
  list := TList<String>.Create;
  list.Add('First'); list.Add('Second');
  list.Exchange(0, 1);
  WriteLn(list[0]);
  WriteLn(list[1]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["Second", "First"]);
}

#[test]
fn test_tlist_move_element() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.Add(10); list.Add(20); list.Add(30);
  list.Move(0, 2);
  WriteLn(list[0]);
  WriteLn(list[2]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["20", "10"]);
}

#[test]
fn test_tlist_first_and_last_properties() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<String>;
begin
  list := TList<String>.Create;
  list.Add('Head'); list.Add('Middle'); list.Add('Tail');
  WriteLn(list.First);
  WriteLn(list.Last);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["Head", "Tail"]);
}

#[test]
fn test_tlist_enum_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TStatus = (stIdle, stRunning, stDone);
var list: TList<TStatus>;
begin
  list := TList<TStatus>.Create;
  list.Add(stIdle); list.Add(stDone);
  WriteLn(Ord(list[0]));
  WriteLn(Ord(list[1]));
  list.Free;
end.
"#);
    assert_eq!(out, vec!["0", "2"]);
}

#[test]
fn test_tlist_addrange() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>;
begin
  list := TList<Integer>.Create;
  list.AddRange([100, 200, 300]);
  WriteLn(list.Count);
  WriteLn(list[1]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["3", "200"]);
}

#[test]
fn test_tlist_lastindexof() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<String>;
begin
  list := TList<String>.Create;
  list.Add('x'); list.Add('y'); list.Add('x');
  WriteLn(list.LastIndexOf('x'));
  list.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_tlist_real_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Real>;
begin
  list := TList<Real>.Create;
  list.Add(1.5); list.Add(2.5);
  WriteLn(list[0] + list[1]);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_tlist_extract_element() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var list: TList<Integer>; val: Integer;
begin
  list := TList<Integer>.Create;
  list.Add(777);
  val := list.Extract(777);
  WriteLn(val);
  WriteLn(list.Count);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["777", "0"]);
}
