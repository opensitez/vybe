// vybe-test: pascal/pascal_linked_list_custom_nodes/test_singly_linked_list_append
// origin: languages/pascal/tests/pascal/test_pascal_linked_list_custom_nodes.rs
program Test;
{$mode delphi}
uses SysUtils;
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
