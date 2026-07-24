use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 49: Binary Search Trees & Tree Traversals
// ═══════════════════════════════════════════════════════════

#[test]
fn test_bst_insertion_and_inorder_traversal() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record
       Key: Integer;
       Left, Right: PNode;
     end;

procedure Insert(var root: PNode; key: Integer);
begin
  if root = nil then
  begin
    New(root);
    root^.Key := key;
    root^.Left := nil; root^.Right := nil;
  end else if key < root^.Key then Insert(root^.Left, key)
  else Insert(root^.Right, key);
end;

procedure InOrder(root: PNode);
begin
  if root = nil then Exit;
  InOrder(root^.Left);
  WriteLn(root^.Key);
  InOrder(root^.Right);
end;

procedure FreeTree(root: PNode);
begin
  if root = nil then Exit;
  FreeTree(root^.Left);
  FreeTree(root^.Right);
  Dispose(root);
end;

var root: PNode;
begin
  root := nil;
  Insert(root, 50); Insert(root, 30); Insert(root, 70);
  InOrder(root);
  FreeTree(root);
end.
"#);
    assert_eq!(out, vec!["30", "50", "70"]);
}

#[test]
fn test_bst_search_key() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

procedure Insert(var root: PNode; key: Integer);
begin
  if root = nil then
  begin
    New(root); root^.Key := key; root^.Left := nil; root^.Right := nil;
  end else if key < root^.Key then Insert(root^.Left, key)
  else Insert(root^.Right, key);
end;

function Search(root: PNode; target: Integer): Boolean;
begin
  if root = nil then Exit(False);
  if root^.Key = target then Exit(True);
  if target < root^.Key then Result := Search(root^.Left, target)
  else Result := Search(root^.Right, target);
end;

procedure FreeTree(root: PNode);
begin
  if root = nil then Exit;
  FreeTree(root^.Left); FreeTree(root^.Right); Dispose(root);
end;

var root: PNode;
begin
  root := nil;
  Insert(root, 40); Insert(root, 20); Insert(root, 60);
  WriteLn(Search(root, 20));
  WriteLn(Search(root, 99));
  FreeTree(root);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_preorder_traversal() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

procedure PreOrder(root: PNode);
begin
  if root = nil then Exit;
  WriteLn(root^.Key);
  PreOrder(root^.Left);
  PreOrder(root^.Right);
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 1;
  New(l); l^.Key := 2; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 3; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  PreOrder(r);
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_postorder_traversal() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

procedure PostOrder(root: PNode);
begin
  if root = nil then Exit;
  PostOrder(root^.Left);
  PostOrder(root^.Right);
  WriteLn(root^.Key);
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 1;
  New(l); l^.Key := 2; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 3; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  PostOrder(r);
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["2", "3", "1"]);
}

#[test]
fn test_bst_min_and_max_key() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

procedure Insert(var root: PNode; key: Integer);
begin
  if root = nil then
  begin
    New(root); root^.Key := key; root^.Left := nil; root^.Right := nil;
  end else if key < root^.Key then Insert(root^.Left, key)
  else Insert(root^.Right, key);
end;

function FindMin(root: PNode): Integer;
begin
  while root^.Left <> nil do root := root^.Left;
  Result := root^.Key;
end;

function FindMax(root: PNode): Integer;
begin
  while root^.Right <> nil do root := root^.Right;
  Result := root^.Key;
end;

procedure FreeTree(root: PNode);
begin
  if root = nil then Exit;
  FreeTree(root^.Left); FreeTree(root^.Right); Dispose(root);
end;

var root: PNode;
begin
  root := nil;
  Insert(root, 50); Insert(root, 10); Insert(root, 90);
  WriteLn(FindMin(root));
  WriteLn(FindMax(root));
  FreeTree(root);
end.
"#);
    assert_eq!(out, vec!["10", "90"]);
}

#[test]
fn test_tree_height_calculation() {
    let out = run_pascal(r#"
program Test;
uses Math;
type PNode = ^TNode;
     TNode = record Left, Right: PNode; end;

function GetHeight(root: PNode): Integer;
begin
  if root = nil then Exit(0);
  Result := 1 + Max(GetHeight(root^.Left), GetHeight(root^.Right));
end;

var r, l, l1: PNode;
begin
  New(r); New(l); New(l1);
  r^.Left := l; r^.Right := nil;
  l^.Left := l1; l^.Right := nil;
  l1^.Left := nil; l1^.Right := nil;
  WriteLn(GetHeight(r));
  Dispose(l1); Dispose(l); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_tree_node_count() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Left, Right: PNode; end;

function CountNodes(root: PNode): Integer;
begin
  if root = nil then Exit(0);
  Result := 1 + CountNodes(root^.Left) + CountNodes(root^.Right);
end;

var r, l, g: PNode;
begin
  New(r); New(l); New(g);
  r^.Left := l; r^.Right := g;
  l^.Left := nil; l^.Right := nil;
  g^.Left := nil; g^.Right := nil;
  WriteLn(CountNodes(r));
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_tree_key_sum() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

function SumKeys(root: PNode): Integer;
begin
  if root = nil then Exit(0);
  Result := root^.Key + SumKeys(root^.Left) + SumKeys(root^.Right);
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 10;
  New(l); l^.Key := 20; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 30; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  WriteLn(SumKeys(r));
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_tree_mirror_inplace() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

procedure MirrorTree(root: PNode);
var temp: PNode;
begin
  if root = nil then Exit;
  temp := root^.Left;
  root^.Left := root^.Right;
  root^.Right := temp;
  MirrorTree(root^.Left);
  MirrorTree(root^.Right);
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 1;
  New(l); l^.Key := 2; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 3; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  MirrorTree(r);
  WriteLn(r^.Left^.Key);
  WriteLn(r^.Right^.Key);
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn test_tree_with_string_key() {
    let out = run_pascal(r#"
program Test;
type PStrNode = ^TStrNode;
     TStrNode = record Key: String; Left, Right: PStrNode; end;

procedure InsertStr(var root: PStrNode; key: String);
begin
  if root = nil then
  begin
    New(root); root^.Key := key; root^.Left := nil; root^.Right := nil;
  end else if key < root^.Key then InsertStr(root^.Left, key)
  else InsertStr(root^.Right, key);
end;

procedure InOrderStr(root: PStrNode);
begin
  if root = nil then Exit;
  InOrderStr(root^.Left);
  WriteLn(root^.Key);
  InOrderStr(root^.Right);
end;

procedure FreeStrTree(root: PStrNode);
begin
  if root = nil then Exit;
  FreeStrTree(root^.Left); FreeStrTree(root^.Right); Dispose(root);
end;

var root: PStrNode;
begin
  root := nil;
  InsertStr(root, 'Banana'); InsertStr(root, 'Apple'); InsertStr(root, 'Cherry');
  InOrderStr(root);
  FreeStrTree(root);
end.
"#);
    assert_eq!(out, vec!["Apple", "Banana", "Cherry"]);
}

#[test]
fn test_is_valid_bst_check() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

function IsBST(root: PNode; minVal, maxVal: Integer): Boolean;
begin
  if root = nil then Exit(True);
  if (root^.Key <= minVal) or (root^.Key >= maxVal) then Exit(False);
  Result := IsBST(root^.Left, minVal, root^.Key) and IsBST(root^.Right, root^.Key, maxVal);
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 20;
  New(l); l^.Key := 10; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 30; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  WriteLn(IsBST(r, -1000, 1000));
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tree_leaf_count() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Left, Right: PNode; end;

function CountLeaves(root: PNode): Integer;
begin
  if root = nil then Exit(0);
  if (root^.Left = nil) and (root^.Right = nil) then Exit(1);
  Result := CountLeaves(root^.Left) + CountLeaves(root^.Right);
end;

var r, l, g: PNode;
begin
  New(r); New(l); New(g);
  r^.Left := l; r^.Right := g;
  l^.Left := nil; l^.Right := nil;
  g^.Left := nil; g^.Right := nil;
  WriteLn(CountLeaves(r));
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_tree_node_record_payload() {
    let out = run_pascal(r#"
program Test;
type TData = record Name: String; Val: Integer; end;
type PNode = ^TNode;
     TNode = record Payload: TData; Left, Right: PNode; end;

var n: PNode;
begin
  New(n);
  n^.Payload.Name := 'RootPayload'; n^.Payload.Val := 777;
  n^.Left := nil; n^.Right := nil;
  WriteLn(n^.Payload.Name);
  WriteLn(n^.Payload.Val);
  Dispose(n);
end.
"#);
    assert_eq!(out, vec!["RootPayload", "777"]);
}

#[test]
fn test_tree_level_order_traversal_mock() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

procedure PrintLevel1(root: PNode);
begin
  if root <> nil then WriteLn(root^.Key);
end;
procedure PrintLevel2(root: PNode);
begin
  if root^.Left <> nil then WriteLn(root^.Left^.Key);
  if root^.Right <> nil then WriteLn(root^.Right^.Key);
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 1;
  New(l); l^.Key := 2; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 3; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  PrintLevel1(r);
  PrintLevel2(r);
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_lowest_common_ancestor_bst() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

function LCA(root: PNode; n1, n2: Integer): Integer;
begin
  if root = nil then Exit(-1);
  if (n1 < root^.Key) and (n2 < root^.Key) then Result := LCA(root^.Left, n1, n2)
  else if (n1 > root^.Key) and (n2 > root^.Key) then Result := LCA(root^.Right, n1, n2)
  else Result := root^.Key;
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 20;
  New(l); l^.Key := 10; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 30; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  WriteLn(LCA(r, 10, 30));
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_bst_delete_leaf_node() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

procedure DeleteLeaf(var root: PNode; target: Integer);
begin
  if root = nil then Exit;
  if root^.Key = target then
  begin
    Dispose(root);
    root := nil;
  end else if target < root^.Key then DeleteLeaf(root^.Left, target)
  else DeleteLeaf(root^.Right, target);
end;

var r, l: PNode;
begin
  New(r); r^.Key := 10;
  New(l); l^.Key := 5; l^.Left := nil; l^.Right := nil;
  r^.Left := l; r^.Right := nil;
  DeleteLeaf(r, 5);
  WriteLn(r^.Left = nil);
  Dispose(r);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tree_path_sum_check() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

function HasPathSum(root: PNode; sum: Integer): Boolean;
begin
  if root = nil then Exit(False);
  if (root^.Left = nil) and (root^.Right = nil) then Exit(sum = root^.Key);
  Result := HasPathSum(root^.Left, sum - root^.Key) or HasPathSum(root^.Right, sum - root^.Key);
end;

var r, l: PNode;
begin
  New(r); r^.Key := 10;
  New(l); l^.Key := 5; l^.Left := nil; l^.Right := nil;
  r^.Left := l; r^.Right := nil;
  WriteLn(HasPathSum(r, 15));
  WriteLn(HasPathSum(r, 99));
  Dispose(l); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_tree_symmetric_check() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

function IsMirror(t1, t2: PNode): Boolean;
begin
  if (t1 = nil) and (t2 = nil) then Exit(True);
  if (t1 = nil) or (t2 = nil) then Exit(False);
  Result := (t1^.Key = t2^.Key) and IsMirror(t1^.Left, t2^.Right) and IsMirror(t1^.Right, t2^.Left);
end;

var r, l, g: PNode;
begin
  New(r); r^.Key := 1;
  New(l); l^.Key := 2; l^.Left := nil; l^.Right := nil;
  New(g); g^.Key := 2; g^.Left := nil; g^.Right := nil;
  r^.Left := l; r^.Right := g;
  WriteLn(IsMirror(r^.Left, r^.Right));
  Dispose(l); Dispose(g); Dispose(r);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tree_copy_structure() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;

function CopyTree(root: PNode): PNode;
begin
  if root = nil then Exit(nil);
  New(Result);
  Result^.Key := root^.Key;
  Result^.Left := CopyTree(root^.Left);
  Result^.Right := CopyTree(root^.Right);
end;

procedure FreeTree(root: PNode);
begin
  if root = nil then Exit;
  FreeTree(root^.Left); FreeTree(root^.Right); Dispose(root);
end;

var orig, copy: PNode;
begin
  New(orig); orig^.Key := 99; orig^.Left := nil; orig^.Right := nil;
  copy := CopyTree(orig);
  WriteLn(copy^.Key);
  FreeTree(orig); FreeTree(copy);
end.
"#);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_tree_single_node_root() {
    let out = run_pascal(r#"
program Test;
type PNode = ^TNode;
     TNode = record Key: Integer; Left, Right: PNode; end;
var root: PNode;
begin
  New(root); root^.Key := 42; root^.Left := nil; root^.Right := nil;
  WriteLn(root^.Key);
  Dispose(root);
end.
"#);
    assert_eq!(out, vec!["42"]);
}
