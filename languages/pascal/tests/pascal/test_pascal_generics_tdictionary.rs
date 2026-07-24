use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 42: Generic Maps & Key-Value Dictionaries (TDictionary<K,V>)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tdictionary_add_and_indexer_lookup() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('one', 1);
  dict.Add('two', 2);
  WriteLn(dict['one']);
  WriteLn(dict['two']);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn test_tdictionary_trygetvalue_success() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, String>;
    val: String;
begin
  dict := TDictionary<String, String>.Create;
  dict.Add('host', 'localhost');
  if dict.TryGetValue('host', val) then
    WriteLn(val);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["localhost"]);
}

#[test]
fn test_tdictionary_trygetvalue_missing_key() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, String>;
    val: String;
begin
  dict := TDictionary<String, String>.Create;
  WriteLn(dict.TryGetValue('port', val));
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_tdictionary_containskey_and_containsvalue() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<Integer, String>;
begin
  dict := TDictionary<Integer, String>.Create;
  dict.Add(200, 'OK');
  dict.Add(404, 'NotFound');
  WriteLn(dict.ContainsKey(200));
  WriteLn(dict.ContainsKey(500));
  WriteLn(dict.ContainsValue('OK'));
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["True", "False", "True"]);
}

#[test]
fn test_tdictionary_addorsetvalue_update() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.AddOrSetValue('count', 10);
  WriteLn(dict['count']);
  dict.AddOrSetValue('count', 20);
  WriteLn(dict['count']);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_tdictionary_remove_key() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('a', 1); dict.Add('b', 2);
  dict.Remove('a');
  WriteLn(dict.Count);
  WriteLn(dict.ContainsKey('a'));
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["1", "False"]);
}

#[test]
fn test_tdictionary_clear_all() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('k1', 10); dict.Add('k2', 20);
  dict.Clear;
  WriteLn(dict.Count);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tdictionary_keys_iteration() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
    k: String;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('A', 1); dict.Add('B', 2);
  for k in dict.Keys do
    WriteLn(k);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn test_tdictionary_values_iteration() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
    v, sum: Integer;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('x', 10); dict.Add('y', 20); dict.Add('z', 30);
  sum := 0;
  for v in dict.Values do
    sum := sum + v;
  WriteLn(sum);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_tdictionary_integer_keys() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<Integer, String>;
begin
  dict := TDictionary<Integer, String>.Create;
  dict.Add(1, 'Jan'); dict.Add(2, 'Feb');
  WriteLn(dict[1] + '-' + dict[2]);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["Jan-Feb"]);
}

#[test]
fn test_tdictionary_record_values() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TPoint = record X, Y: Integer; end;
var dict: TDictionary<String, TPoint>; pt: TPoint;
begin
  dict := TDictionary<String, TPoint>.Create;
  pt.X := 15; pt.Y := 30;
  dict.Add('start', pt);
  WriteLn(dict['start'].X);
  WriteLn(dict['start'].Y);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["15", "30"]);
}

#[test]
fn test_tdictionary_enum_keys() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TStatus = (stLow, stMed, stHigh);
var dict: TDictionary<TStatus, String>;
begin
  dict := TDictionary<TStatus, String>.Create;
  dict.Add(stLow, 'Green'); dict.Add(stHigh, 'Red');
  WriteLn(dict[stLow]);
  WriteLn(dict[stHigh]);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["Green", "Red"]);
}

#[test]
fn test_tdictionary_real_values() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Real>;
begin
  dict := TDictionary<String, Real>.Create;
  dict.Add('pi', 3.14159);
  dict.Add('e', 2.71828);
  WriteLn(dict['pi']);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn test_tdictionary_boolean_values() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Boolean>;
begin
  dict := TDictionary<String, Boolean>.Create;
  dict.Add('enabled', True);
  dict.Add('debug', False);
  WriteLn(dict['enabled']);
  WriteLn(dict['debug']);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_tdictionary_extractpair() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
    pair: TPair<String, Integer>;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('item', 500);
  pair := dict.ExtractPair('item');
  WriteLn(pair.Key);
  WriteLn(pair.Value);
  WriteLn(dict.Count);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["item", "500", "0"]);
}

#[test]
fn test_tdictionary_toarray() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
    arr: TArray<TPair<String, Integer>>;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('A', 1); dict.Add('B', 2);
  arr := dict.ToArray;
  WriteLn(Length(arr));
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_tdictionary_key_case_sensitivity() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, String>;
begin
  dict := TDictionary<String, String>.Create;
  dict.Add('key', 'lower');
  dict.Add('KEY', 'upper');
  WriteLn(dict['key']);
  WriteLn(dict['KEY']);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["lower", "upper"]);
}

#[test]
fn test_tdictionary_count_tracking() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<Integer, Integer>; i: Integer;
begin
  dict := TDictionary<Integer, Integer>.Create;
  for i := 1 to 5 do dict.Add(i, i * 10);
  WriteLn(dict.Count);
  dict.Remove(3);
  WriteLn(dict.Count);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["5", "4"]);
}

#[test]
fn test_tdictionary_pair_iteration() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, Integer>;
    pair: TPair<String, Integer>;
begin
  dict := TDictionary<String, Integer>.Create;
  dict.Add('X', 100);
  for pair in dict do
  begin
    WriteLn(pair.Key);
    WriteLn(pair.Value);
  end;
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["X", "100"]);
}

#[test]
fn test_tdictionary_reinsertion_after_remove() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var dict: TDictionary<String, String>;
begin
  dict := TDictionary<String, String>.Create;
  dict.Add('temp', 'v1');
  dict.Remove('temp');
  dict.Add('temp', 'v2');
  WriteLn(dict['temp']);
  dict.Free;
end.
"#);
    assert_eq!(out, vec!["v2"]);
}
