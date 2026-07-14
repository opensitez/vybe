program LinkedListDemo;

type
  PNode = ^TNode;
  TNode = record
    Data: Integer;
    Next: PNode;
  end;

  TLinkedList = class
  private
    FHead: PNode;
    FCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddFirst(Value: Integer);
    procedure AddLast(Value: Integer);
    procedure Remove(Value: Integer);
    function Contains(Value: Integer): Boolean;
    procedure Print;
    property Count: Integer read FCount;
  end;

constructor TLinkedList.Create;
begin
  FHead := nil;
  FCount := 0;
end;

destructor TLinkedList.Destroy;
var
  Current, NextNode: PNode;
begin
  Current := FHead;
  while Current <> nil do
  begin
    NextNode := Current^.Next;
    Dispose(Current);
    Current := NextNode;
  end;
end;

procedure TLinkedList.AddFirst(Value: Integer);
var
  NewNode: PNode;
begin
  New(NewNode);
  NewNode^.Data := Value;
  NewNode^.Next := FHead;
  FHead := NewNode;
  FCount := FCount + 1;
end;

procedure TLinkedList.AddLast(Value: Integer);
var
  NewNode, Current: PNode;
begin
  New(NewNode);
  NewNode^.Data := Value;
  NewNode^.Next := nil;
  if FHead = nil then
    FHead := NewNode
  else
  begin
    Current := FHead;
    while Current^.Next <> nil do
      Current := Current^.Next;
    Current^.Next := NewNode;
  end;
  FCount := FCount + 1;
end;

procedure TLinkedList.Remove(Value: Integer);
var
  Current, Prev: PNode;
begin
  if FHead = nil then Exit;
  if FHead^.Data = Value then
  begin
    Current := FHead;
    FHead := FHead^.Next;
    Dispose(Current);
    FCount := FCount - 1;
    Exit;
  end;
  Prev := FHead;
  Current := FHead^.Next;
  while Current <> nil do
  begin
    if Current^.Data = Value then
    begin
      Prev^.Next := Current^.Next;
      Dispose(Current);
      FCount := FCount - 1;
      Exit;
    end;
    Prev := Current;
    Current := Current^.Next;
  end;
end;

function TLinkedList.Contains(Value: Integer): Boolean;
var
  Current: PNode;
begin
  Current := FHead;
  while Current <> nil do
  begin
    if Current^.Data = Value then
    begin
      Result := True;
      Exit;
    end;
    Current := Current^.Next;
  end;
  Result := False;
end;

procedure TLinkedList.Print;
var
  Current: PNode;
begin
  Current := FHead;
  while Current <> nil do
  begin
    Write(Current^.Data, ' ');
    Current := Current^.Next;
  end;
  Writeln;
end;

var
  List: TLinkedList;
begin
  List := TLinkedList.Create;
  List.AddLast(10);
  List.AddLast(20);
  List.AddFirst(5);
  List.AddLast(30);

  Writeln('List contents:');
  List.Print;
  Writeln('Count: ', List.Count);
  Writeln('Contains 20? ', List.Contains(20));

  List.Remove(20);
  Writeln('After removing 20:');
  List.Print;

  List.Free;
end.
