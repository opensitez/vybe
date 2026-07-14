program PointersDemo;

type
  PInteger = ^Integer;
  PNode = ^TNode;
  TNode = record
    Value: Integer;
    Next: PNode;
  end;

procedure PrintList(Head: PNode);
var
  Current: PNode;
begin
  Current := Head;
  while Current <> nil do
  begin
    Write(Current^.Value, ' ');
    Current := Current^.Next;
  end;
  Writeln;
end;

var
  Ptr: PInteger;
  A: Integer;
  Head, Second, Third: PNode;
begin
  A := 42;
  Ptr := @A;
  Writeln('Ptr^ = ', Ptr^);
  Ptr^ := 100;
  Writeln('A = ', A);

  New(Head);
  Head^.Value := 1;
  New(Second);
  Second^.Value := 2;
  New(Third);
  Third^.Value := 3;

  Head^.Next := Second;
  Second^.Next := Third;
  Third^.Next := nil;

  Writeln('Linked list:');
  PrintList(Head);

  Dispose(Third);
  Dispose(Second);
  Dispose(Head);
end.
