use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 45: TStringList Properties, Formatting & Key-Value Parsing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tstringlist_add_and_strings_indexer() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('Line1');
  sl.Add('Line2');
  WriteLn(sl.Count);
  WriteLn(sl[0]);
  WriteLn(sl[1]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2", "Line1", "Line2"]);
}

#[test]
fn test_tstringlist_text_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Text := 'Alpha' + #13#10 + 'Beta' + #13#10 + 'Gamma';
  WriteLn(sl.Count);
  WriteLn(sl[1]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3", "Beta"]);
}

#[test]
fn test_tstringlist_commatext_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.CommaText := 'apple,banana,"cherry, ripe"';
  WriteLn(sl.Count);
  WriteLn(sl[0]);
  WriteLn(sl[2]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3", "apple", "cherry, ripe"]);
}

#[test]
fn test_tstringlist_delimitedtext_custom_delimiter() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Delimiter := '|';
  sl.StrictDelimiter := True;
  sl.DelimitedText := 'one|two|three';
  WriteLn(sl.Count);
  WriteLn(sl[0]);
  WriteLn(sl[2]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["3", "one", "three"]);
}

#[test]
fn test_tstringlist_key_value_values_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('host=localhost');
  sl.Add('port=8080');
  WriteLn(sl.Values['host']);
  WriteLn(sl.Values['port']);
  sl.Values['port'] := '9090';
  WriteLn(sl.Values['port']);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["localhost", "8080", "9090"]);
}

#[test]
fn test_tstringlist_names_indexer() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('User=Alice');
  sl.Add('Role=Admin');
  WriteLn(sl.Names[0]);
  WriteLn(sl.Names[1]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["User", "Role"]);
}

#[test]
fn test_tstringlist_sorted_and_find() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList; idx: Integer;
begin
  sl := TStringList.Create;
  sl.Sorted := True;
  sl.Add('Zebra'); sl.Add('Apple'); sl.Add('Monkey');
  WriteLn(sl[0]);
  if sl.Find('Monkey', idx) then
    WriteLn(idx);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Apple", "1"]);
}

#[test]
fn test_tstringlist_duplicates_ignore() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Sorted := True;
  sl.Duplicates := dupIgnore;
  sl.Add('Item');
  sl.Add('Item');
  WriteLn(sl.Count);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_tstringlist_addobject_and_objects() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.AddObject('First', TObject(100));
  sl.AddObject('Second', TObject(200));
  WriteLn(NativeInt(sl.Objects[0]));
  WriteLn(NativeInt(sl.Objects[1]));
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_tstringlist_casesensitive_search() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.CaseSensitive := True;
  sl.Add('Pascal');
  WriteLn(sl.IndexOf('pascal'));
  WriteLn(sl.IndexOf('Pascal'));
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["-1", "0"]);
}

#[test]
fn test_tstringlist_clear_all_lines() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('Line1'); sl.Add('Line2');
  sl.Clear;
  WriteLn(sl.Count);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tstringlist_delete_and_insert() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('A'); sl.Add('C');
  sl.Insert(1, 'B');
  WriteLn(sl[1]);
  sl.Delete(0);
  WriteLn(sl[0]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["B", "B"]);
}

#[test]
fn test_tstringlist_exchange_lines() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('Top'); sl.Add('Bottom');
  sl.Exchange(0, 1);
  WriteLn(sl[0]);
  WriteLn(sl[1]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Bottom", "Top"]);
}

#[test]
fn test_tstringlist_indexofname_lookup() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('mode=test');
  sl.Add('env=production');
  WriteLn(sl.IndexOfName('env'));
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_tstringlist_valuefromindex_property() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Add('key1=val1');
  sl.Add('key2=val2');
  WriteLn(sl.ValueFromIndex[1]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["val2"]);
}

#[test]
fn test_tstringlist_ownsobjects_automatic_freeing() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
type TItemObj = class
  public Name: String;
  constructor Create(N: String); destructor Destroy; override;
end;
constructor TItemObj.Create(N: String); begin Name := N; end;
destructor TItemObj.Destroy; begin WriteLn('FreedObj:' + Name); inherited Destroy; end;

var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.OwnsObjects := True;
  sl.AddObject('Obj1', TItemObj.Create('Item1'));
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["FreedObj:Item1"]);
}

#[test]
fn test_tstringlist_namevalueseparator_custom() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.NameValueSeparator := ':';
  sl.Add('Host:127.0.0.1');
  WriteLn(sl.Names[0]);
  WriteLn(sl.Values['Host']);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Host", "127.0.0.1"]);
}

#[test]
fn test_tstringlist_for_in_loop_iteration() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList; line: String;
begin
  sl := TStringList.Create;
  sl.Add('Row1'); sl.Add('Row2');
  for line in sl do
    WriteLn(line);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Row1", "Row2"]);
}

#[test]
fn test_tstringlist_sort_custom_case_insensitive() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.CaseSensitive := False;
  sl.Add('b'); sl.Add('A'); sl.Add('c');
  sl.Sort;
  WriteLn(sl[0]);
  WriteLn(sl[1]);
  WriteLn(sl[2]);
  sl.Free;
end.
"#,
    );
    assert_eq!(out, vec!["A", "b", "c"]);
}

#[test]
fn test_tstringlist_assign_copy() {
    let out = run_pascal(
        r#"
program Test;
uses Classes;
var sl1, sl2: TStringList;
begin
  sl1 := TStringList.Create;
  sl1.Add('CopyLine1'); sl1.Add('CopyLine2');
  sl2 := TStringList.Create;
  sl2.Assign(sl1);
  WriteLn(sl2.Count);
  WriteLn(sl2[1]);
  sl1.Free; sl2.Free;
end.
"#,
    );
    assert_eq!(out, vec!["2", "CopyLine2"]);
}
