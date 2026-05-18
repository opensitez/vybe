/// Tests for generic container simulations in Pascal/Delphi:
/// Generic pair operations, generic stack/queue patterns,
/// generic search and transform functions.

use super::helpers::run_pascal;

// ===================================================================
// GENERIC PAIR OPERATIONS
// ===================================================================

#[test] fn generic_pair_swap() {
    assert_eq!(run_pascal(r#"program T;
type
  TPair<T> = class
  public
    First: T;
    Second: T;
    procedure Swap;
  end;
procedure TPair<T>.Swap;
var tmp: T;
begin
  tmp := First;
  First := Second;
  Second := tmp;
end;
var p: TPair<Integer>;
begin
  p := TPair<Integer>.Create;
  p.First := 10;
  p.Second := 20;
  p.Swap;
  WriteLn(p.First);
  WriteLn(p.Second);
  p.Free;
end."#), &["20", "10"]);
}

#[test] fn generic_pair_string() {
    assert_eq!(run_pascal(r#"program T;
type
  TPair<T> = class
  public
    First: T;
    Second: T;
  end;
var p: TPair<String>;
begin
  p := TPair<String>.Create;
  p.First := 'hello';
  p.Second := 'world';
  WriteLn(p.First + ' ' + p.Second);
  p.Free;
end."#), &["hello world"]);
}

// ===================================================================
// GENERIC STACK SIMULATION
// ===================================================================

#[test] fn generic_stack_push_pop() {
    assert_eq!(run_pascal(r#"program T;
type
  TStack<T> = class
  private
    FItems: array of T;
    FCount: Integer;
  public
    procedure Push(item: T);
    function Pop: T;
    function Peek: T;
    function IsEmpty: Boolean;
    function Count: Integer;
  end;
procedure TStack<T>.Push(item: T);
begin
  SetLength(FItems, FCount + 1);
  FItems[FCount] := item;
  FCount := FCount + 1;
end;
function TStack<T>.Pop: T;
begin
  FCount := FCount - 1;
  Result := FItems[FCount];
end;
function TStack<T>.Peek: T;
begin
  Result := FItems[FCount - 1];
end;
function TStack<T>.IsEmpty: Boolean;
begin
  Result := FCount = 0;
end;
function TStack<T>.Count: Integer;
begin
  Result := FCount;
end;
var s: TStack<Integer>;
begin
  s := TStack<Integer>.Create;
  s.Push(1);
  s.Push(2);
  s.Push(3);
  WriteLn(s.Count);
  WriteLn(s.Peek);
  WriteLn(s.Pop);
  WriteLn(s.Count);
  s.Free;
end."#), &["3", "3", "3", "2"]);
}

#[test] fn generic_stack_string() {
    assert_eq!(run_pascal(r#"program T;
type
  TStack<T> = class
  private
    FItems: array of T;
    FCount: Integer;
  public
    procedure Push(item: T);
    function Pop: T;
  end;
procedure TStack<T>.Push(item: T);
begin
  SetLength(FItems, FCount + 1);
  FItems[FCount] := item;
  FCount := FCount + 1;
end;
function TStack<T>.Pop: T;
begin
  FCount := FCount - 1;
  Result := FItems[FCount];
end;
var s: TStack<String>;
begin
  s := TStack<String>.Create;
  s.Push('first');
  s.Push('second');
  s.Push('third');
  WriteLn(s.Pop);
  WriteLn(s.Pop);
  s.Free;
end."#), &["third", "second"]);
}

// ===================================================================
// GENERIC QUEUE SIMULATION
// ===================================================================

#[test] fn generic_queue_enqueue_dequeue() {
    assert_eq!(run_pascal(r#"program T;
type
  TQueue<T> = class
  private
    FItems: array of T;
    FHead: Integer;
    FTail: Integer;
  public
    constructor Create;
    procedure Enqueue(item: T);
    function Dequeue: T;
    function IsEmpty: Boolean;
  end;
constructor TQueue<T>.Create;
begin
  inherited;
  FHead := 0;
  FTail := 0;
end;
procedure TQueue<T>.Enqueue(item: T);
begin
  SetLength(FItems, FTail + 1);
  FItems[FTail] := item;
  FTail := FTail + 1;
end;
function TQueue<T>.Dequeue: T;
begin
  Result := FItems[FHead];
  FHead := FHead + 1;
end;
function TQueue<T>.IsEmpty: Boolean;
begin
  Result := FHead >= FTail;
end;
var q: TQueue<Integer>;
begin
  q := TQueue<Integer>.Create;
  q.Enqueue(10);
  q.Enqueue(20);
  q.Enqueue(30);
  WriteLn(q.Dequeue);
  WriteLn(q.Dequeue);
  WriteLn(q.IsEmpty);
  q.Free;
end."#), &["10", "20", "false"]);
}

// ===================================================================
// GENERIC FUNCTION
// ===================================================================

#[test] fn generic_max_function() {
    assert_eq!(run_pascal(r#"program T;
function GenMax<T>(a, b: T): T;
begin
  if a > b then Result := a else Result := b;
end;
begin
  WriteLn(GenMax<Integer>(3, 7));
  WriteLn(GenMax<Integer>(10, 5));
end."#), &["7", "10"]);
}

#[test] fn generic_swap_function() {
    assert_eq!(run_pascal(r#"program T;
procedure GenSwap<T>(var a, b: T);
var tmp: T;
begin
  tmp := a;
  a := b;
  b := tmp;
end;
var x, y: Integer;
begin
  x := 100;
  y := 200;
  GenSwap<Integer>(x, y);
  WriteLn(x);
  WriteLn(y);
end."#), &["200", "100"]);
}

// ===================================================================
// GENERIC CONTAINER WITH FIND
// ===================================================================

#[test] fn generic_list_contains() {
    assert_eq!(run_pascal(r#"program T;
type
  TList<T> = class
  private
    FItems: array of T;
    FCount: Integer;
  public
    procedure Add(item: T);
    function Contains(item: T): Boolean;
  end;
procedure TList<T>.Add(item: T);
begin
  SetLength(FItems, FCount + 1);
  FItems[FCount] := item;
  FCount := FCount + 1;
end;
function TList<T>.Contains(item: T): Boolean;
var i: Integer;
begin
  Result := False;
  for i := 0 to FCount - 1 do
    if FItems[i] = item then
    begin
      Result := True;
      Break;
    end;
end;
var lst: TList<Integer>;
begin
  lst := TList<Integer>.Create;
  lst.Add(10);
  lst.Add(20);
  lst.Add(30);
  WriteLn(lst.Contains(20));
  WriteLn(lst.Contains(99));
  lst.Free;
end."#), &["true", "false"]);
}

// ===================================================================
// GENERIC PAIR COMPARISON
// ===================================================================

#[test] fn generic_pair_equal() {
    assert_eq!(run_pascal(r#"program T;
type
  TPair<T> = class
  public
    A: T;
    B: T;
    function Equal: Boolean;
  end;
function TPair<T>.Equal: Boolean;
begin
  Result := A = B;
end;
var p: TPair<Integer>;
begin
  p := TPair<Integer>.Create;
  p.A := 5;
  p.B := 5;
  WriteLn(p.Equal);
  p.B := 6;
  WriteLn(p.Equal);
  p.Free;
end."#), &["true", "false"]);
}

// ===================================================================
// GENERIC CLASS WITH COUNT
// ===================================================================

#[test] fn generic_bag_count() {
    assert_eq!(run_pascal(r#"program T;
type
  TBag<T> = class
  private
    FItems: array of T;
    FCount: Integer;
  public
    procedure Add(item: T);
    function Count: Integer;
    function ItemAt(i: Integer): T;
  end;
procedure TBag<T>.Add(item: T);
begin
  SetLength(FItems, FCount + 1);
  FItems[FCount] := item;
  FCount := FCount + 1;
end;
function TBag<T>.Count: Integer;
begin
  Result := FCount;
end;
function TBag<T>.ItemAt(i: Integer): T;
begin
  Result := FItems[i];
end;
var bag: TBag<String>;
begin
  bag := TBag<String>.Create;
  bag.Add('apple');
  bag.Add('banana');
  bag.Add('cherry');
  WriteLn(bag.Count);
  WriteLn(bag.ItemAt(1));
  bag.Free;
end."#), &["3", "banana"]);
}

// ===================================================================
// GENERIC PAIR USED IN FUNCTION
// ===================================================================

#[test] fn generic_pair_as_function_return() {
    assert_eq!(run_pascal(r#"program T;
type
  TIntPair = class
  public
    First: Integer;
    Second: Integer;
  end;
function MinMax(arr: array of Integer): TIntPair;
var i: Integer;
begin
  Result := TIntPair.Create;
  Result.First := arr[0];
  Result.Second := arr[0];
  for i := 1 to Length(arr) - 1 do
  begin
    if arr[i] < Result.First then Result.First := arr[i];
    if arr[i] > Result.Second then Result.Second := arr[i];
  end;
end;
var mm: TIntPair;
    data: array of Integer;
begin
  SetLength(data, 5);
  data[0] := 5; data[1] := 2; data[2] := 8; data[3] := 1; data[4] := 9;
  mm := MinMax(data);
  WriteLn(mm.First);
  WriteLn(mm.Second);
  mm.Free;
end."#), &["1", "9"]);
}

// ===================================================================
// GENERIC FUNCTION WITH STRING
// ===================================================================

#[test] fn generic_function_identity() {
    assert_eq!(run_pascal(r#"program T;
function Identity<T>(value: T): T;
begin
  Result := value;
end;
begin
  WriteLn(Identity<Integer>(42));
  WriteLn(Identity<String>('hello'));
  WriteLn(Identity<Boolean>(True));
end."#), &["42", "hello", "true"]);
}

// ===================================================================
// GENERIC CLASS INHERITANCE
// ===================================================================

#[test] fn generic_base_class() {
    assert_eq!(run_pascal(r#"program T;
type
  TContainer<T> = class
  protected
    FValue: T;
  public
    procedure SetValue(v: T);
    function GetValue: T;
  end;
  TNumberBox<T> = class(TContainer<T>)
  public
    function IsZero: Boolean;
  end;
procedure TContainer<T>.SetValue(v: T);
begin
  FValue := v;
end;
function TContainer<T>.GetValue: T;
begin
  Result := FValue;
end;
function TNumberBox<T>.IsZero: Boolean;
begin
  Result := FValue = 0;
end;
var box: TNumberBox<Integer>;
begin
  box := TNumberBox<Integer>.Create;
  box.SetValue(0);
  WriteLn(box.IsZero);
  box.SetValue(5);
  WriteLn(box.IsZero);
  WriteLn(box.GetValue);
  box.Free;
end."#), &["true", "false", "5"]);
}

// ===================================================================
// GENERIC MAP/TRANSFORM
// ===================================================================

#[test] fn generic_transform_array() {
    assert_eq!(run_pascal(r#"program T;
type
  TTransformer = class
    class function DoubleAll(arr: array of Integer): TArray<Integer>;
  end;
class function TTransformer.DoubleAll(arr: array of Integer): TArray<Integer>;
var i: Integer;
begin
  SetLength(Result, Length(arr));
  for i := 0 to Length(arr) - 1 do
    Result[i] := arr[i] * 2;
end;
var input, output: array of Integer;
    i: Integer;
begin
  SetLength(input, 3);
  input[0] := 1; input[1] := 2; input[2] := 3;
  output := TTransformer.DoubleAll(input);
  for i := 0 to Length(output) - 1 do
    WriteLn(output[i]);
end."#), &["2", "4", "6"]);
}

// ===================================================================
// GENERIC PAIR IN LIST
// ===================================================================

#[test] fn generic_key_value_pair() {
    assert_eq!(run_pascal(r#"program T;
type
  TKeyValue<K, V> = class
  public
    Key: K;
    Value: V;
    constructor Create(aKey: K; aValue: V);
    function ToString: String; override;
  end;
constructor TKeyValue<K, V>.Create(aKey: K; aValue: V);
begin
  inherited Create;
  Key := aKey;
  Value := aValue;
end;
function TKeyValue<K, V>.ToString: String;
begin
  Result := IntToStr(Key) + '=' + Value;
end;
var kv: TKeyValue<Integer, String>;
begin
  kv := TKeyValue<Integer, String>.Create(1, 'one');
  WriteLn(kv.ToString);
  kv.Free;
  kv := TKeyValue<Integer, String>.Create(2, 'two');
  WriteLn(kv.ToString);
  kv.Free;
end."#), &["1=one", "2=two"]);
}
