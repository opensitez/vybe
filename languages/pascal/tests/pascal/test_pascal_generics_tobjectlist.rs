use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 44: Owned Object Containers (TObjectList<T>)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tobjectlist_ownsobjects_automatic_destruction() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TSampleItem = class
  public ID: Integer;
  constructor Create(AID: Integer); destructor Destroy; override;
end;
constructor TSampleItem.Create(AID: Integer); begin ID := AID; end;
destructor TSampleItem.Destroy; begin WriteLn('FreedItem:' + ID.ToString); inherited Destroy; end;

var list: TObjectList<TSampleItem>;
begin
  list := TObjectList<TSampleItem>.Create(True);
  list.Add(TSampleItem.Create(1));
  list.Add(TSampleItem.Create(2));
  list.Free;
end.
"#);
    assert_eq!(out, vec!["FreedItem:1", "FreedItem:2"]);
}

#[test]
fn test_tobjectlist_delete_frees_owned_object() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TWidget = class
  public Tag: String;
  constructor Create(T: String); destructor Destroy; override;
end;
constructor TWidget.Create(T: String); begin Tag := T; end;
destructor TWidget.Destroy; begin WriteLn('WidgetDestroyed:' + Tag); inherited Destroy; end;

var list: TObjectList<TWidget>;
begin
  list := TObjectList<TWidget>.Create(True);
  list.Add(TWidget.Create('W1'));
  list.Add(TWidget.Create('W2'));
  list.Delete(0);
  WriteLn('AfterDelete');
  list.Free;
end.
"#);
    assert_eq!(out, vec!["WidgetDestroyed:W1", "AfterDelete", "WidgetDestroyed:W2"]);
}

#[test]
fn test_tobjectlist_clear_frees_all_owned_objects() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TNode = class
  destructor Destroy; override;
end;
destructor TNode.Destroy; begin WriteLn('NodeCleared'); inherited Destroy; end;

var list: TObjectList<TNode>;
begin
  list := TObjectList<TNode>.Create(True);
  list.Add(TNode.Create);
  list.Add(TNode.Create);
  list.Clear;
  WriteLn(list.Count);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["NodeCleared", "NodeCleared", "0"]);
}

#[test]
fn test_tobjectlist_ownsobjects_false_does_not_free() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TExternalObj = class
  public Name: String;
  constructor Create(N: String);
end;
constructor TExternalObj.Create(N: String); begin Name := N; end;

var list: TObjectList<TExternalObj>; obj: TExternalObj;
begin
  obj := TExternalObj.Create('External');
  list := TObjectList<TExternalObj>.Create(False);
  list.Add(obj);
  list.Clear;
  WriteLn(obj.Name);
  list.Free;
  obj.Free;
end.
"#);
    assert_eq!(out, vec!["External"]);
}

#[test]
fn test_tobjectlist_extract_without_freeing() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  public Code: Integer;
  constructor Create(C: Integer); destructor Destroy; override;
end;
constructor TItem.Create(C: Integer); begin Code := C; end;
destructor TItem.Destroy; begin WriteLn('ItemFreed:' + Code.ToString); inherited Destroy; end;

var list: TObjectList<TItem>; extracted: TItem;
begin
  list := TObjectList<TItem>.Create(True);
  list.Add(TItem.Create(100));
  extracted := list.Extract(list[0]);
  WriteLn(list.Count);
  WriteLn(extracted.Code);
  list.Free;
  extracted.Free;
end.
"#);
    assert_eq!(out, vec!["0", "100", "ItemFreed:100"]);
}

#[test]
fn test_tobjectlist_replace_item_frees_previous() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TData = class
  public LabelStr: String;
  constructor Create(L: String); destructor Destroy; override;
end;
constructor TData.Create(L: String); begin LabelStr := L; end;
destructor TData.Destroy; begin WriteLn('ReplacedFreed:' + LabelStr); inherited Destroy; end;

var list: TObjectList<TData>;
begin
  list := TObjectList<TData>.Create(True);
  list.Add(TData.Create('Old'));
  list[0] := TData.Create('New');
  WriteLn(list[0].LabelStr);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["ReplacedFreed:Old", "New", "ReplacedFreed:New"]);
}

#[test]
fn test_tobjectlist_first_and_last_accessors() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TElement = class
  public Name: String;
  constructor Create(N: String);
end;
constructor TElement.Create(N: String); begin Name := N; end;

var list: TObjectList<TElement>;
begin
  list := TObjectList<TElement>.Create(True);
  list.Add(TElement.Create('FirstEl'));
  list.Add(TElement.Create('LastEl'));
  WriteLn(list.First.Name);
  WriteLn(list.Last.Name);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["FirstEl", "LastEl"]);
}

#[test]
fn test_tobjectlist_contains_and_indexof() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TObj = class end;
var list: TObjectList<TObj>; o1, o2: TObj;
begin
  list := TObjectList<TObj>.Create(True);
  o1 := TObj.Create; o2 := TObj.Create;
  list.Add(o1);
  WriteLn(list.Contains(o1));
  WriteLn(list.Contains(o2));
  WriteLn(list.IndexOf(o1));
  list.Free;
  o2.Free;
end.
"#);
    assert_eq!(out, vec!["True", "False", "0"]);
}

#[test]
fn test_tobjectlist_for_in_loop() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  public Val: Integer;
  constructor Create(V: Integer);
end;
constructor TItem.Create(V: Integer); begin Val := V; end;

var list: TObjectList<TItem>; item: TItem; sum: Integer;
begin
  list := TObjectList<TItem>.Create(True);
  list.Add(TItem.Create(10)); list.Add(TItem.Create(20));
  sum := 0;
  for item in list do
    sum := sum + item.Val;
  WriteLn(sum);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_tobjectlist_extractat_index() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  public Val: Integer;
  constructor Create(V: Integer); destructor Destroy; override;
end;
constructor TItem.Create(V: Integer); begin Val := V; end;
destructor TItem.Destroy; begin WriteLn('ItemFreed:' + Val.ToString); inherited Destroy; end;

var list: TObjectList<TItem>; extracted: TItem;
begin
  list := TObjectList<TItem>.Create(True);
  list.Add(TItem.Create(10)); list.Add(TItem.Create(20));
  extracted := list.ExtractAt(0);
  WriteLn(extracted.Val);
  WriteLn(list.Count);
  extracted.Free;
  list.Free;
end.
"#);
    assert_eq!(out, vec!["10", "1", "ItemFreed:10", "ItemFreed:20"]);
}

#[test]
fn test_tobjectlist_remove_frees_object() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  destructor Destroy; override;
end;
destructor TItem.Destroy; begin WriteLn('RemovedAndFreed'); inherited Destroy; end;

var list: TObjectList<TItem>; item: TItem;
begin
  list := TObjectList<TItem>.Create(True);
  item := TItem.Create;
  list.Add(item);
  list.Remove(item);
  WriteLn(list.Count);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["RemovedAndFreed", "0"]);
}

#[test]
fn test_tobjectlist_sort_with_custom_comparer() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections, Generics.Defaults;
type TScoreObj = class
  public Score: Integer;
  constructor Create(S: Integer);
end;
constructor TScoreObj.Create(S: Integer); begin Score := S; end;

var list: TObjectList<TScoreObj>;
begin
  list := TObjectList<TScoreObj>.Create(True);
  list.Add(TScoreObj.Create(50));
  list.Add(TScoreObj.Create(10));
  list.Sort(TComparer<TScoreObj>.Construct(
    function(const L, R: TScoreObj): Integer
    begin
      Result := L.Score - R.Score;
    end));
  WriteLn(list[0].Score);
  WriteLn(list[1].Score);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["10", "50"]);
}

#[test]
fn test_tobjectlist_toarray_of_objects() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class end;
var list: TObjectList<TItem>; arr: TArray<TItem>;
begin
  list := TObjectList<TItem>.Create(True);
  list.Add(TItem.Create); list.Add(TItem.Create);
  arr := list.ToArray;
  WriteLn(Length(arr));
  list.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_nested_tobjectlist_hierarchy() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TChild = class
  destructor Destroy; override;
end;
type TParent = class
  public Children: TObjectList<TChild>;
  constructor Create; destructor Destroy; override;
end;
destructor TChild.Destroy; begin WriteLn('ChildFreed'); inherited Destroy; end;
constructor TParent.Create; begin Children := TObjectList<TChild>.Create(True); end;
destructor TParent.Destroy; begin Children.Free; WriteLn('ParentFreed'); inherited Destroy; end;

var list: TObjectList<TParent>; p: TParent;
begin
  list := TObjectList<TParent>.Create(True);
  p := TParent.Create;
  p.Children.Add(TChild.Create);
  list.Add(p);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["ChildFreed", "ParentFreed"]);
}

#[test]
fn test_tobjectlist_exchange_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  public Name: String;
  constructor Create(N: String);
end;
constructor TItem.Create(N: String); begin Name := N; end;

var list: TObjectList<TItem>;
begin
  list := TObjectList<TItem>.Create(True);
  list.Add(TItem.Create('First')); list.Add(TItem.Create('Second'));
  list.Exchange(0, 1);
  WriteLn(list[0].Name);
  WriteLn(list[1].Name);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["Second", "First"]);
}

#[test]
fn test_tobjectlist_reverse_order() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  public Val: Integer;
  constructor Create(V: Integer);
end;
constructor TItem.Create(V: Integer); begin Val := V; end;

var list: TObjectList<TItem>;
begin
  list := TObjectList<TItem>.Create(True);
  list.Add(TItem.Create(1)); list.Add(TItem.Create(2));
  list.Reverse;
  WriteLn(list[0].Val);
  WriteLn(list[1].Val);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn test_tobjectlist_trimexcess() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class end;
var list: TObjectList<TItem>;
begin
  list := TObjectList<TItem>.Create(True);
  list.Capacity := 50;
  list.Add(TItem.Create);
  list.TrimExcess;
  WriteLn(list.Capacity);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_tobjectlist_addrange_objects() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  public Code: Integer;
  constructor Create(C: Integer);
end;
constructor TItem.Create(C: Integer); begin Code := C; end;

var list: TObjectList<TItem>;
begin
  list := TObjectList<TItem>.Create(True);
  list.AddRange([TItem.Create(10), TItem.Create(20)]);
  WriteLn(list.Count);
  WriteLn(list[0].Code);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["2", "10"]);
}

#[test]
fn test_tobjectlist_ownsobjects_toggle_runtime() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class
  destructor Destroy; override;
end;
destructor TItem.Destroy; begin WriteLn('ToggledFreed'); inherited Destroy; end;

var list: TObjectList<TItem>; item: TItem;
begin
  list := TObjectList<TItem>.Create(False);
  item := TItem.Create;
  list.Add(item);
  list.OwnsObjects := True;
  list.Free;
end.
"#);
    assert_eq!(out, vec!["ToggledFreed"]);
}

#[test]
fn test_tobjectlist_empty_list_free() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TItem = class end;
var list: TObjectList<TItem>;
begin
  list := TObjectList<TItem>.Create(True);
  WriteLn(list.Count);
  list.Free;
end.
"#);
    assert_eq!(out, vec!["0"]);
}
