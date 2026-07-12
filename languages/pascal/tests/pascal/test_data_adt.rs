/// Tests for abstract data type implementations using classes in Pascal/Delphi:
/// Stack, Queue, and Deque built on top of dynamic arrays, linked list nodes,
/// and priority queue patterns — going beyond what test_data_structures.rs covers.
use super::helpers::run_pascal;

// ===================================================================
// INTEGER STACK CLASS
// ===================================================================

#[test]
fn int_stack_lifo_order() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TIntStack = class
  private
    FData: array of Integer;
    FTop: Integer;
  public
    constructor Create;
    procedure Push(v: Integer);
    function Pop: Integer;
    function Top: Integer;
    function Size: Integer;
    function IsEmpty: Boolean;
  end;
constructor TIntStack.Create;
begin
  inherited;
  FTop := 0;
end;
procedure TIntStack.Push(v: Integer);
begin
  SetLength(FData, FTop + 1);
  FData[FTop] := v;
  FTop := FTop + 1;
end;
function TIntStack.Pop: Integer;
begin
  FTop := FTop - 1;
  Result := FData[FTop];
end;
function TIntStack.Top: Integer;
begin
  Result := FData[FTop - 1];
end;
function TIntStack.Size: Integer;
begin
  Result := FTop;
end;
function TIntStack.IsEmpty: Boolean;
begin
  Result := FTop = 0;
end;
var s: TIntStack;
begin
  s := TIntStack.Create;
  s.Push(5);
  s.Push(10);
  s.Push(15);
  WriteLn(s.Size);
  WriteLn(s.Top);
  WriteLn(s.Pop);
  WriteLn(s.Pop);
  WriteLn(s.IsEmpty);
  s.Free;
end."#
        ),
        &["3", "15", "15", "10", "false"]
    );
}

#[test]
fn stack_reverse_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TIntStack = class
  private
    FData: array of Integer;
    FTop: Integer;
  public
    constructor Create;
    procedure Push(v: Integer);
    function Pop: Integer;
    function IsEmpty: Boolean;
  end;
constructor TIntStack.Create;
begin
  inherited;
  FTop := 0;
end;
procedure TIntStack.Push(v: Integer);
begin
  SetLength(FData, FTop + 1);
  FData[FTop] := v;
  Inc(FTop);
end;
function TIntStack.Pop: Integer;
begin
  Dec(FTop);
  Result := FData[FTop];
end;
function TIntStack.IsEmpty: Boolean;
begin
  Result := FTop = 0;
end;
var s: TIntStack;
    i: Integer;
begin
  s := TIntStack.Create;
  for i := 1 to 5 do
    s.Push(i);
  while not s.IsEmpty do
    Write(IntToStr(s.Pop) + ' ');
  WriteLn('');
  s.Free;
end."#
        ),
        &["5 4 3 2 1 "]
    );
}

// ===================================================================
// INTEGER QUEUE CLASS
// ===================================================================

#[test]
fn int_queue_fifo_order() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TIntQueue = class
  private
    FData: array of Integer;
    FHead: Integer;
    FTail: Integer;
  public
    constructor Create;
    procedure Enqueue(v: Integer);
    function Dequeue: Integer;
    function Front: Integer;
    function Size: Integer;
    function IsEmpty: Boolean;
  end;
constructor TIntQueue.Create;
begin
  inherited;
  FHead := 0;
  FTail := 0;
end;
procedure TIntQueue.Enqueue(v: Integer);
begin
  SetLength(FData, FTail + 1);
  FData[FTail] := v;
  Inc(FTail);
end;
function TIntQueue.Dequeue: Integer;
begin
  Result := FData[FHead];
  Inc(FHead);
end;
function TIntQueue.Front: Integer;
begin
  Result := FData[FHead];
end;
function TIntQueue.Size: Integer;
begin
  Result := FTail - FHead;
end;
function TIntQueue.IsEmpty: Boolean;
begin
  Result := FHead >= FTail;
end;
var q: TIntQueue;
begin
  q := TIntQueue.Create;
  q.Enqueue(100);
  q.Enqueue(200);
  q.Enqueue(300);
  WriteLn(q.Size);
  WriteLn(q.Front);
  WriteLn(q.Dequeue);
  WriteLn(q.Dequeue);
  WriteLn(q.Size);
  q.Free;
end."#
        ),
        &["3", "100", "100", "200", "1"]
    );
}

#[test]
fn queue_bfs_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TIntQueue = class
  private
    FData: array of Integer;
    FHead, FTail: Integer;
  public
    constructor Create;
    procedure Enqueue(v: Integer);
    function Dequeue: Integer;
    function IsEmpty: Boolean;
  end;
constructor TIntQueue.Create;
begin
  inherited;
  FHead := 0;
  FTail := 0;
end;
procedure TIntQueue.Enqueue(v: Integer);
begin
  SetLength(FData, FTail + 1);
  FData[FTail] := v;
  Inc(FTail);
end;
function TIntQueue.Dequeue: Integer;
begin
  Result := FData[FHead];
  Inc(FHead);
end;
function TIntQueue.IsEmpty: Boolean;
begin
  Result := FHead >= FTail;
end;
var q: TIntQueue;
    total: Integer;
begin
  q := TIntQueue.Create;
  q.Enqueue(1);
  q.Enqueue(2);
  q.Enqueue(4);
  q.Enqueue(8);
  total := 0;
  while not q.IsEmpty do
    total := total + q.Dequeue;
  WriteLn(total);
  q.Free;
end."#
        ),
        &["15"]
    );
}

// ===================================================================
// DEQUE (DOUBLE-ENDED QUEUE) CLASS
// ===================================================================

#[test]
fn deque_push_front_back() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TDeque = class
  private
    FData: array of Integer;
    FCount: Integer;
  public
    constructor Create;
    procedure PushBack(v: Integer);
    procedure PushFront(v: Integer);
    function PopBack: Integer;
    function PopFront: Integer;
    function Count: Integer;
  end;
constructor TDeque.Create;
begin
  inherited;
  FCount := 0;
end;
procedure TDeque.PushBack(v: Integer);
begin
  SetLength(FData, FCount + 1);
  FData[FCount] := v;
  Inc(FCount);
end;
procedure TDeque.PushFront(v: Integer);
var i: Integer;
begin
  SetLength(FData, FCount + 1);
  for i := FCount downto 1 do
    FData[i] := FData[i - 1];
  FData[0] := v;
  Inc(FCount);
end;
function TDeque.PopBack: Integer;
begin
  Dec(FCount);
  Result := FData[FCount];
end;
function TDeque.PopFront: Integer;
var i: Integer;
begin
  Result := FData[0];
  for i := 0 to FCount - 2 do
    FData[i] := FData[i + 1];
  Dec(FCount);
end;
function TDeque.Count: Integer;
begin
  Result := FCount;
end;
var d: TDeque;
begin
  d := TDeque.Create;
  d.PushBack(2);
  d.PushFront(1);
  d.PushBack(3);
  WriteLn(d.Count);
  WriteLn(d.PopFront);
  WriteLn(d.PopBack);
  d.Free;
end."#
        ),
        &["3", "1", "3"]
    );
}

// ===================================================================
// LINKED LIST NODE
// ===================================================================

#[test]
fn linked_list_build() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TNode = class
  public
    Value: Integer;
    Next: TNode;
    constructor Create(v: Integer);
  end;
constructor TNode.Create(v: Integer);
begin
  inherited Create;
  Value := v;
  Next := nil;
end;
var head, curr: TNode;
    sum: Integer;
begin
  head := TNode.Create(1);
  head.Next := TNode.Create(2);
  head.Next.Next := TNode.Create(3);
  sum := 0;
  curr := head;
  while curr <> nil do
  begin
    sum := sum + curr.Value;
    curr := curr.Next;
  end;
  WriteLn(sum);
end."#
        ),
        &["6"]
    );
}

#[test]
fn linked_list_length() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TNode = class
  public
    Value: Integer;
    Next: TNode;
  end;
function ListLength(head: TNode): Integer;
var curr: TNode;
begin
  Result := 0;
  curr := head;
  while curr <> nil do
  begin
    Inc(Result);
    curr := curr.Next;
  end;
end;
var head: TNode;
begin
  head := TNode.Create;
  head.Value := 10;
  head.Next := TNode.Create;
  head.Next.Value := 20;
  head.Next.Next := TNode.Create;
  head.Next.Next.Value := 30;
  head.Next.Next.Next := nil;
  WriteLn(ListLength(head));
end."#
        ),
        &["3"]
    );
}

// ===================================================================
// PRIORITY QUEUE (min-heap simulation)
// ===================================================================

#[test]
fn priority_queue_simple() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TPriorityQueue = class
  private
    FData: array of Integer;
    FCount: Integer;
    function MinIndex: Integer;
  public
    constructor Create;
    procedure Insert(v: Integer);
    function ExtractMin: Integer;
    function IsEmpty: Boolean;
  end;
constructor TPriorityQueue.Create;
begin
  inherited;
  FCount := 0;
end;
function TPriorityQueue.MinIndex: Integer;
var i, m: Integer;
begin
  m := 0;
  for i := 1 to FCount - 1 do
    if FData[i] < FData[m] then m := i;
  Result := m;
end;
procedure TPriorityQueue.Insert(v: Integer);
begin
  SetLength(FData, FCount + 1);
  FData[FCount] := v;
  Inc(FCount);
end;
function TPriorityQueue.ExtractMin: Integer;
var idx: Integer;
begin
  idx := MinIndex;
  Result := FData[idx];
  FData[idx] := FData[FCount - 1];
  Dec(FCount);
end;
function TPriorityQueue.IsEmpty: Boolean;
begin
  Result := FCount = 0;
end;
var pq: TPriorityQueue;
begin
  pq := TPriorityQueue.Create;
  pq.Insert(30);
  pq.Insert(10);
  pq.Insert(20);
  WriteLn(pq.ExtractMin);
  WriteLn(pq.ExtractMin);
  WriteLn(pq.ExtractMin);
  pq.Free;
end."#
        ),
        &["10", "20", "30"]
    );
}

// ===================================================================
// CIRCULAR BUFFER
// ===================================================================

#[test]
fn circular_buffer_wrap() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TCircBuf = class
  private
    FData: array[0..4] of Integer;
    FHead: Integer;
    FTail: Integer;
    FSize: Integer;
    FCapacity: Integer;
  public
    constructor Create;
    procedure Push(v: Integer);
    function Pop: Integer;
    function Count: Integer;
  end;
constructor TCircBuf.Create;
begin
  inherited;
  FHead := 0;
  FTail := 0;
  FSize := 0;
  FCapacity := 5;
end;
procedure TCircBuf.Push(v: Integer);
begin
  FData[FTail] := v;
  FTail := (FTail + 1) mod FCapacity;
  Inc(FSize);
end;
function TCircBuf.Pop: Integer;
begin
  Result := FData[FHead];
  FHead := (FHead + 1) mod FCapacity;
  Dec(FSize);
end;
function TCircBuf.Count: Integer;
begin
  Result := FSize;
end;
var cb: TCircBuf;
begin
  cb := TCircBuf.Create;
  cb.Push(1);
  cb.Push(2);
  cb.Push(3);
  WriteLn(cb.Pop);
  cb.Push(4);
  WriteLn(cb.Count);
  WriteLn(cb.Pop);
  cb.Free;
end."#
        ),
        &["1", "3", "2"]
    );
}

// ===================================================================
// STRING STACK
// ===================================================================

#[test]
fn string_stack_operations() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TStrStack = class
  private
    FData: array of String;
    FTop: Integer;
  public
    constructor Create;
    procedure Push(s: String);
    function Pop: String;
    function Peek: String;
    function IsEmpty: Boolean;
  end;
constructor TStrStack.Create;
begin
  inherited;
  FTop := 0;
end;
procedure TStrStack.Push(s: String);
begin
  SetLength(FData, FTop + 1);
  FData[FTop] := s;
  Inc(FTop);
end;
function TStrStack.Pop: String;
begin
  Dec(FTop);
  Result := FData[FTop];
end;
function TStrStack.Peek: String;
begin
  Result := FData[FTop - 1];
end;
function TStrStack.IsEmpty: Boolean;
begin
  Result := FTop = 0;
end;
var ss: TStrStack;
begin
  ss := TStrStack.Create;
  ss.Push('first');
  ss.Push('second');
  ss.Push('third');
  WriteLn(ss.Peek);
  WriteLn(ss.Pop);
  WriteLn(ss.Pop);
  ss.Free;
end."#
        ),
        &["third", "third", "second"]
    );
}

// ===================================================================
// ASSOCIATIVE MAP SIMULATION (key-value pairs array)
// ===================================================================

#[test]
fn simple_map_lookup() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMapEntry = record
    Key: String;
    Value: Integer;
  end;
  TSimpleMap = class
  private
    FEntries: array of TMapEntry;
    FCount: Integer;
  public
    procedure Put(key: String; value: Integer);
    function Get(key: String; def: Integer): Integer;
  end;
procedure TSimpleMap.Put(key: String; value: Integer);
var i: Integer;
begin
  for i := 0 to FCount - 1 do
    if FEntries[i].Key = key then
    begin
      FEntries[i].Value := value;
      Exit;
    end;
  SetLength(FEntries, FCount + 1);
  FEntries[FCount].Key := key;
  FEntries[FCount].Value := value;
  Inc(FCount);
end;
function TSimpleMap.Get(key: String; def: Integer): Integer;
var i: Integer;
begin
  for i := 0 to FCount - 1 do
    if FEntries[i].Key = key then
    begin
      Result := FEntries[i].Value;
      Exit;
    end;
  Result := def;
end;
var m: TSimpleMap;
begin
  m := TSimpleMap.Create;
  m.Put('a', 1);
  m.Put('b', 2);
  m.Put('c', 3);
  WriteLn(m.Get('b', 0));
  WriteLn(m.Get('z', -1));
  m.Put('b', 99);
  WriteLn(m.Get('b', 0));
  m.Free;
end."#
        ),
        &["2", "-1", "99"]
    );
}

// ===================================================================
// MULTISET (bag with counts)
// ===================================================================

#[test]
fn multiset_add_count() {
    assert_eq!(
        run_pascal(
            r#"program T;
type
  TMultiSet = class
  private
    FKeys: array of String;
    FCounts: array of Integer;
    FSize: Integer;
  public
    procedure Add(item: String);
    function CountOf(item: String): Integer;
  end;
procedure TMultiSet.Add(item: String);
var i: Integer;
begin
  for i := 0 to FSize - 1 do
    if FKeys[i] = item then
    begin
      Inc(FCounts[i]);
      Exit;
    end;
  SetLength(FKeys, FSize + 1);
  SetLength(FCounts, FSize + 1);
  FKeys[FSize] := item;
  FCounts[FSize] := 1;
  Inc(FSize);
end;
function TMultiSet.CountOf(item: String): Integer;
var i: Integer;
begin
  for i := 0 to FSize - 1 do
    if FKeys[i] = item then
    begin
      Result := FCounts[i];
      Exit;
    end;
  Result := 0;
end;
var ms: TMultiSet;
begin
  ms := TMultiSet.Create;
  ms.Add('apple');
  ms.Add('banana');
  ms.Add('apple');
  ms.Add('cherry');
  ms.Add('apple');
  WriteLn(ms.CountOf('apple'));
  WriteLn(ms.CountOf('banana'));
  WriteLn(ms.CountOf('grape'));
  ms.Free;
end."#
        ),
        &["3", "1", "0"]
    );
}
