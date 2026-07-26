use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 48: Singly & Doubly Linked List Structures
// ═══════════════════════════════════════════════════════════

#[test]
fn test_singly_linked_list_prepend_and_traverse() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record
       Val: Integer;
       Next: PNode;
     end;

procedure Prepend(var head: PNode; val: Integer);
var newHead: PNode;
begin
  New(newHead);
  newHead^.Val := val;
  newHead^.Next := head;
  head := newHead;
end;

var head, curr, temp: PNode;
begin
  head := nil;
  Prepend(head, 10);
  Prepend(head, 20);
  Prepend(head, 30);
  curr := head;
  while curr <> nil do
  begin
    WriteLn(curr^.Val);
    curr := curr^.Next;
  end;
  while head <> nil do
  begin
    temp := head; head := head^.Next; Dispose(temp);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["30", "20", "10"]);
}

#[test]
fn test_singly_linked_list_append() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;

procedure Append(var head: PNode; val: Integer);
var newNode, curr: PNode;
begin
  New(newNode); newNode^.Val := val; newNode^.Next := nil;
  if head = nil then head := newNode
  else begin
    curr := head;
    while curr^.Next <> nil do curr := curr^.Next;
    curr^.Next := newNode;
  end;
end;

var head, curr, temp: PNode;
begin
  head := nil;
  Append(head, 100); Append(head, 200);
  curr := head;
  while curr <> nil do begin WriteLn(curr^.Val); curr := curr^.Next; end;
  while head <> nil do begin temp := head; head := head^.Next; Dispose(temp); end;
end.
"#,
    );
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_doubly_linked_list_forward_and_backward() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record
       Val: Integer;
       Prev, Next: PNode;
     end;

var n1, n2: PNode;
begin
  New(n1); New(n2);
  n1^.Val := 1; n1^.Prev := nil; n1^.Next := n2;
  n2^.Val := 2; n2^.Prev := n1;  n2^.Next := nil;

  WriteLn(n1^.Val);
  WriteLn(n1^.Next^.Val);
  WriteLn(n2^.Prev^.Val);

  Dispose(n1); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["1", "2", "1"]);
}

#[test]
fn test_linked_list_delete_node() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;

procedure DeleteVal(var head: PNode; target: Integer);
var curr, prev, temp: PNode;
begin
  curr := head; prev := nil;
  while (curr <> nil) and (curr^.Val <> target) do
  begin
    prev := curr; curr := curr^.Next;
  end;
  if curr <> nil then
  begin
    if prev = nil then head := curr^.Next
    else prev^.Next := curr^.Next;
    Dispose(curr);
  end;
end;

var head, curr, temp: PNode;
begin
  New(head); head^.Val := 10;
  New(head^.Next); head^.Next^.Val := 20; head^.Next^.Next := nil;
  DeleteVal(head, 10);
  WriteLn(head^.Val);
  Dispose(head);
end.
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_singly_linked_list_inplace_reverse() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;

procedure ReverseList(var head: PNode);
var prev, curr, nextNode: PNode;
begin
  prev := nil; curr := head;
  while curr <> nil do
  begin
    nextNode := curr^.Next;
    curr^.Next := prev;
    prev := curr;
    curr := nextNode;
  end;
  head := prev;
end;

var head, curr, temp: PNode;
begin
  New(head); head^.Val := 1;
  New(head^.Next); head^.Next^.Val := 2;
  New(head^.Next^.Next); head^.Next^.Next^.Val := 3; head^.Next^.Next^.Next := nil;
  ReverseList(head);
  curr := head;
  while curr <> nil do begin WriteLn(curr^.Val); curr := curr^.Next; end;
  while head <> nil do begin temp := head; head := head^.Next; Dispose(temp); end;
end.
"#,
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn test_linked_list_count_nodes() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;
function CountNodes(head: PNode): Integer;
begin
  Result := 0;
  while head <> nil do begin Inc(Result); head := head^.Next; end;
end;
var n1, n2, n3: PNode;
begin
  New(n1); New(n2); New(n3);
  n1^.Next := n2; n2^.Next := n3; n3^.Next := nil;
  WriteLn(CountNodes(n1));
  Dispose(n1); Dispose(n2); Dispose(n3);
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_linked_list_sum_values() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;
var n1, n2: PNode; sum: Integer;
begin
  New(n1); New(n2);
  n1^.Val := 15; n1^.Next := n2;
  n2^.Val := 35; n2^.Next := nil;
  sum := n1^.Val + n2^.Val;
  WriteLn(sum);
  Dispose(n1); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_circular_linked_list_traversal() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;
var n1, n2: PNode; curr: PNode; count: Integer;
begin
  New(n1); New(n2);
  n1^.Val := 100; n1^.Next := n2;
  n2^.Val := 200; n2^.Next := n1;
  curr := n1; count := 0;
  repeat
    WriteLn(curr^.Val);
    curr := curr^.Next;
    Inc(count);
  until count = 4;
  Dispose(n1); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["100", "200", "100", "200"]);
}

#[test]
fn test_linked_list_with_string_payload() {
    let out = run_pascal(
        r#"
program Test;
type PStrNode = ^TStrNode;
     TStrNode = record Text: String; Next: PStrNode; end;
var n1, n2: PStrNode;
begin
  New(n1); New(n2);
  n1^.Text := 'Head'; n1^.Next := n2;
  n2^.Text := 'Tail'; n2^.Next := nil;
  WriteLn(n1^.Text + ' -> ' + n2^.Text);
  Dispose(n1); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["Head -> Tail"]);
}

#[test]
fn test_linked_list_find_node_by_value() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;
function FindNode(head: PNode; target: Integer): Boolean;
begin
  Result := False;
  while head <> nil do
  begin
    if head^.Val = target then Exit(True);
    head := head^.Next;
  end;
end;
var n1, n2: PNode;
begin
  New(n1); New(n2);
  n1^.Val := 5; n1^.Next := n2; n2^.Val := 10; n2^.Next := nil;
  WriteLn(FindNode(n1, 10));
  WriteLn(FindNode(n1, 99));
  Dispose(n1); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_linked_list_queue_adapter() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;
type TQueue = record Head, Tail: PNode; end;

procedure Push(var q: TQueue; v: Integer);
var n: PNode;
begin
  New(n); n^.Val := v; n^.Next := nil;
  if q.Tail = nil then begin q.Head := n; q.Tail := n; end
  else begin q.Tail^.Next := n; q.Tail := n; end;
end;

function Pop(var q: TQueue): Integer;
var temp: PNode;
begin
  Result := q.Head^.Val;
  temp := q.Head; q.Head := q.Head^.Next;
  if q.Head = nil then q.Tail := nil;
  Dispose(temp);
end;

var q: TQueue;
begin
  q.Head := nil; q.Tail := nil;
  Push(q, 10); Push(q, 20);
  WriteLn(Pop(q));
  WriteLn(Pop(q));
end.
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_linked_list_stack_adapter() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;
type TStack = record Top: PNode; end;

procedure Push(var s: TStack; v: Integer);
var n: PNode;
begin
  New(n); n^.Val := v; n^.Next := s.Top; s.Top := n;
end;

function Pop(var s: TStack): Integer;
var temp: PNode;
begin
  Result := s.Top^.Val;
  temp := s.Top; s.Top := s.Top^.Next;
  Dispose(temp);
end;

var s: TStack;
begin
  s.Top := nil;
  Push(s, 100); Push(s, 200);
  WriteLn(Pop(s));
  WriteLn(Pop(s));
end.
"#,
    );
    assert_eq!(out, vec!["200", "100"]);
}

#[test]
fn test_linked_list_insert_after() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;

var n1, n2, nNew: PNode;
begin
  New(n1); New(n2);
  n1^.Val := 1; n1^.Next := n2; n2^.Val := 3; n2^.Next := nil;
  New(nNew); nNew^.Val := 2;
  nNew^.Next := n1^.Next;
  n1^.Next := nNew;

  WriteLn(n1^.Next^.Val);
  WriteLn(n1^.Next^.Next^.Val);
  Dispose(n1); Dispose(nNew); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn test_linked_list_record_payload() {
    let out = run_pascal(
        r#"
program Test;
type TPoint = record X, Y: Integer; end;
type PRecNode = ^TRecNode;
     TRecNode = record Data: TPoint; Next: PRecNode; end;

var n: PRecNode;
begin
  New(n);
  n^.Data.X := 12; n^.Data.Y := 24; n^.Next := nil;
  WriteLn(n^.Data.X + n^.Data.Y);
  Dispose(n);
end.
"#,
    );
    assert_eq!(out, vec!["36"]);
}

#[test]
fn test_floyd_cycle_detection_no_cycle() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Next: PNode; end;
function HasCycle(head: PNode): Boolean;
var slow, fast: PNode;
begin
  slow := head; fast := head; Result := False;
  while (fast <> nil) and (fast^.Next <> nil) do
  begin
    slow := slow^.Next;
    fast := fast^.Next^.Next;
    if slow = fast then Exit(True);
  end;
end;
var n1, n2: PNode;
begin
  New(n1); New(n2); n1^.Next := n2; n2^.Next := nil;
  WriteLn(HasCycle(n1));
  Dispose(n1); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_floyd_cycle_detection_with_cycle() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Next: PNode; end;
function HasCycle(head: PNode): Boolean;
var slow, fast: PNode;
begin
  slow := head; fast := head; Result := False;
  while (fast <> nil) and (fast^.Next <> nil) do
  begin
    slow := slow^.Next;
    fast := fast^.Next^.Next;
    if slow = fast then Exit(True);
  end;
end;
var n1, n2: PNode;
begin
  New(n1); New(n2); n1^.Next := n2; n2^.Next := n1;
  WriteLn(HasCycle(n1));
  Dispose(n1); Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linked_list_deep_copy() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;

function DeepCopy(head: PNode): PNode;
var newHead, last, n: PNode;
begin
  if head = nil then Exit(nil);
  New(newHead); newHead^.Val := head^.Val; newHead^.Next := nil;
  last := newHead; head := head^.Next;
  while head <> nil do
  begin
    New(n); n^.Val := head^.Val; n^.Next := nil;
    last^.Next := n; last := n; head := head^.Next;
  end;
  Result := newHead;
end;

var orig, copyList: PNode;
begin
  New(orig); orig^.Val := 42; orig^.Next := nil;
  copyList := DeepCopy(orig);
  WriteLn(copyList^.Val);
  Dispose(orig); Dispose(copyList);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_linked_list_middle_element_find() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;
function GetMiddle(head: PNode): Integer;
var slow, fast: PNode;
begin
  slow := head; fast := head;
  while (fast <> nil) and (fast^.Next <> nil) do
  begin
    slow := slow^.Next;
    fast := fast^.Next^.Next;
  end;
  Result := slow^.Val;
end;
var n1, n2, n3: PNode;
begin
  New(n1); New(n2); New(n3);
  n1^.Val := 10; n1^.Next := n2;
  n2^.Val := 20; n2^.Next := n3;
  n3^.Val := 30; n3^.Next := nil;
  WriteLn(GetMiddle(n1));
  Dispose(n1); Dispose(n2); Dispose(n3);
end.
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_linked_list_merge_two_sorted_lists() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Next: PNode; end;

function MergeSorted(l1, l2: PNode): PNode;
var dummy: TNode; tail: PNode;
begin
  dummy.Next := nil; tail := @dummy;
  while (l1 <> nil) and (l2 <> nil) do
  begin
    if l1^.Val <= l2^.Val then begin tail^.Next := l1; l1 := l1^.Next; end
    else begin tail^.Next := l2; l2 := l2^.Next; end;
    tail := tail^.Next;
  end;
  if l1 <> nil then tail^.Next := l1 else tail^.Next := l2;
  Result := dummy.Next;
end;

var a, b, merged: PNode;
begin
  New(a); a^.Val := 10; a^.Next := nil;
  New(b); b^.Val := 20; b^.Next := nil;
  merged := MergeSorted(a, b);
  WriteLn(merged^.Val);
  WriteLn(merged^.Next^.Val);
  Dispose(a); Dispose(b);
end.
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_doubly_linked_list_remove_head() {
    let out = run_pascal(
        r#"
program Test;
type PNode = ^TNode;
     TNode = record Val: Integer; Prev, Next: PNode; end;

var n1, n2: PNode;
begin
  New(n1); New(n2);
  n1^.Val := 5;  n1^.Prev := nil; n1^.Next := n2;
  n2^.Val := 15; n2^.Prev := n1;  n2^.Next := nil;

  n2^.Prev := nil;
  Dispose(n1);
  WriteLn(n2^.Val);
  Dispose(n2);
end.
"#,
    );
    assert_eq!(out, vec!["15"]);
}
