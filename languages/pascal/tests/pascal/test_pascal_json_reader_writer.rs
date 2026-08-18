use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 69: JSON DOM & Serialization (TJSONObject & TJSONArray)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_jsonobject_create_and_addpair_string() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  json.AddPair('title', 'PascalDoc');
  WriteLn(json.Values['title'].Value);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["PascalDoc"]);
}

#[test]
fn test_jsonobject_addpair_number_and_boolean() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  json.AddPair('id', TJSONNumber.Create(101));
  json.AddPair('active', TJSONTrue.Create);
  WriteLn(json.Values['id'].Value);
  WriteLn(json.Values['active'].Value);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["101", "true"]);
}

#[test]
fn test_jsonarray_add_elements() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var arr: TJSONArray;
begin
  arr := TJSONArray.Create;
  arr.Add('Item1');
  arr.Add('Item2');
  WriteLn(arr.Count);
  WriteLn(arr.Items[0].Value);
  WriteLn(arr.Items[1].Value);
  arr.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2", "Item1", "Item2"]);
}

#[test]
fn test_jsonobject_tostring_output() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  json.AddPair('key', 'value');
  WriteLn(json.ToString);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["{\"key\":\"value\"}"]);
}

#[test]
fn test_jsonobject_parsejsonvalue_object() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var val: TJSONValue; json: TJSONObject;
begin
  val := TJSONObject.ParseJSONValue('{"name":"Alice","age":30}');
  json := val as TJSONObject;
  WriteLn(json.Values['name'].Value);
  WriteLn(json.Values['age'].Value);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn test_jsonobject_parsejsonvalue_array() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var val: TJSONValue; arr: TJSONArray;
begin
  val := TJSONObject.ParseJSONValue('["red","green","blue"]');
  arr := val as TJSONArray;
  WriteLn(arr.Count);
  WriteLn(arr.Items[1].Value);
  arr.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3", "green"]);
}

#[test]
fn test_jsonobject_nested_object() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var root, child: TJSONObject;
begin
  root := TJSONObject.Create;
  child := TJSONObject.Create;
  child.AddPair('city', 'Metropolis');
  root.AddPair('address', child);

  WriteLn((root.Values['address'] as TJSONObject).Values['city'].Value);
  root.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Metropolis"]);
}

#[test]
fn test_jsonobject_getvalue_generic_helper() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  json.AddPair('count', TJSONNumber.Create(42));
  WriteLn(json.GetValue<Integer>('count'));
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_jsonobject_jsonnull_handling() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  json.AddPair('data', TJSONNull.Create);
  WriteLn(json.Values['data'] is TJSONNull);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_jsonobject_contains_key_check() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  json.AddPair('status', 'ok');
  WriteLn(json.Values['status'] <> nil);
  WriteLn(json.Values['missing'] = nil);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE"]);
}

#[test]
fn test_jsonobject_pairs_count_and_iteration() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject; pair: TJSONPair;
begin
  json := TJSONObject.Create;
  json.AddPair('a', '1'); json.AddPair('b', '2');
  WriteLn(json.Count);
  for pair in json do
    WriteLn(pair.JsonString.Value + '=' + pair.JsonValue.Value);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2", "a=1", "b=2"]);
}

#[test]
fn test_jsonarray_for_in_loop() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var arr: TJSONArray; val: TJSONValue;
begin
  arr := TJSONArray.Create;
  arr.Add('X'); arr.Add('Y');
  for val in arr do
    WriteLn(val.Value);
  arr.Free;
end.
"#,
    );
    assert_eq!(out, vec!["X", "Y"]);
}

#[test]
fn test_jsonobject_removepair() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject; pair: TJSONPair;
begin
  json := TJSONObject.Create;
  json.AddPair('temp', 'to_remove');
  pair := json.RemovePair('temp');
  WriteLn(json.Count);
  pair.Free;
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_record_to_jsonobject_serialization() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
type TUserRec = record ID: Integer; Name: String; end;
function UserToJSON(const u: TUserRec): TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('id', TJSONNumber.Create(u.ID));
  Result.AddPair('name', u.Name);
end;
var user: TUserRec; json: TJSONObject;
begin
  user.ID := 99; user.Name := 'Bob';
  json := UserToJSON(user);
  WriteLn(json.ToString);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["{\"id\":99,\"name\":\"Bob\"}"]);
}

#[test]
fn test_jsonobject_to_record_deserialization() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
type TUserRec = record ID: Integer; Name: String; end;
function JSONToUser(json: TJSONObject): TUserRec;
begin
  Result.ID := json.GetValue<Integer>('id');
  Result.Name := json.GetValue<String>('name');
end;
var json: TJSONObject; user: TUserRec;
begin
  json := TJSONObject.ParseJSONValue('{"id":88,"name":"Charlie"}') as TJSONObject;
  user := JSONToUser(json);
  WriteLn(user.ID.ToString + ':' + user.Name);
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["88:Charlie"]);
}

#[test]
fn test_jsonobject_clone_deep_copy() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var original, cloneObj: TJSONObject;
begin
  original := TJSONObject.Create;
  original.AddPair('key', 'val');
  cloneObj := original.Clone as TJSONObject;
  WriteLn(cloneObj.Values['key'].Value);
  original.Free; cloneObj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["val"]);
}

#[test]
fn test_jsonobject_format_pretty_print() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  json.AddPair('k', 'v');
  WriteLn(json.Format(2) <> '');
  json.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_jsonarray_nested_objects() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var arr: TJSONArray; o1, o2: TJSONObject;
begin
  arr := TJSONArray.Create;
  o1 := TJSONObject.Create; o1.AddPair('id', TJSONNumber.Create(1));
  o2 := TJSONObject.Create; o2.AddPair('id', TJSONNumber.Create(2));
  arr.Add(o1); arr.Add(o2);

  WriteLn((arr.Items[0] as TJSONObject).Values['id'].Value);
  WriteLn((arr.Items[1] as TJSONObject).Values['id'].Value);
  arr.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn test_jsonobject_parse_invalid_returns_nil() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var val: TJSONValue;
begin
  val := TJSONObject.ParseJSONValue('{invalid_json}');
  WriteLn(val = nil);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_jsonobject_protection_finally() {
    let out = run_pascal(
        r#"
program Test;
uses System.JSON;
var json: TJSONObject;
begin
  json := TJSONObject.Create;
  try
    json.AddPair('status', 'ok');
    WriteLn('JSONCreated');
  finally
    json.Free;
    WriteLn('JSONFreedInFinally');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["JSONCreated", "JSONFreedInFinally"]);
}
