program StackQueueDemo;

type
  TIntStack = class
  private
    FItems: array of Integer;
  public
    procedure Push(Value: Integer);
    function Pop: Integer;
    function Peek: Integer;
    function Count: Integer;
    function IsEmpty: Boolean;
  end;

  TIntQueue = class
  private
    FItems: array of Integer;
  public
    procedure Enqueue(Value: Integer);
    function Dequeue: Integer;
    function Front: Integer;
    function Count: Integer;
    function IsEmpty: Boolean;
  end;

procedure TIntStack.Push(Value: Integer);
begin
  SetLength(FItems, Length(FItems) + 1);
  FItems[High(FItems)] := Value;
end;

function TIntStack.Pop: Integer;
begin
  Result := FItems[High(FItems)];
  SetLength(FItems, Length(FItems) - 1);
end;

function TIntStack.Peek: Integer;
begin
  Result := FItems[High(FItems)];
end;

function TIntStack.Count: Integer;
begin
  Result := Length(FItems);
end;

function TIntStack.IsEmpty: Boolean;
begin
  Result := Count = 0;
end;

procedure TIntQueue.Enqueue(Value: Integer);
begin
  SetLength(FItems, Length(FItems) + 1);
  FItems[High(FItems)] := Value;
end;

function TIntQueue.Dequeue: Integer;
var
  I: Integer;
begin
  Result := FItems[0];
  for I := 0 to High(FItems) - 1 do
    FItems[I] := FItems[I + 1];
  SetLength(FItems, Length(FItems) - 1);
end;

function TIntQueue.Front: Integer;
begin
  Result := FItems[0];
end;

function TIntQueue.Count: Integer;
begin
  Result := Length(FItems);
end;

function TIntQueue.IsEmpty: Boolean;
begin
  Result := Count = 0;
end;

var
  Stack: TIntStack;
  Queue: TIntQueue;
  I: Integer;
begin
  Stack := TIntStack.Create;
  for I := 1 to 5 do
    Stack.Push(I * 10);
  Writeln('Stack peek: ', Stack.Peek);
  while not Stack.IsEmpty do
    Write(Stack.Pop, ' ');
  Writeln;
  Stack.Free;

  Queue := TIntQueue.Create;
  for I := 1 to 5 do
    Queue.Enqueue(I * 10);
  Writeln('Queue front: ', Queue.Front);
  while not Queue.IsEmpty do
    Write(Queue.Dequeue, ' ');
  Writeln;
  Queue.Free;
end.
