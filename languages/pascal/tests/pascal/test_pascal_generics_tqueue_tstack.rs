use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 43: Generic Queues & Stacks (TQueue<T>, TStack<T>)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_tqueue_integer_enqueue_dequeue_fifo() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<Integer>;
begin
  q := TQueue<Integer>.Create;
  q.Enqueue(10);
  q.Enqueue(20);
  q.Enqueue(30);
  WriteLn(q.Dequeue);
  WriteLn(q.Dequeue);
  WriteLn(q.Dequeue);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn test_tqueue_peek_front() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<String>;
begin
  q := TQueue<String>.Create;
  q.Enqueue('First');
  q.Enqueue('Second');
  WriteLn(q.Peek);
  WriteLn(q.Count);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["First", "2"]);
}

#[test]
fn test_tstack_integer_push_pop_lifo() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<Integer>;
begin
  s := TStack<Integer>.Create;
  s.Push(10);
  s.Push(20);
  s.Push(30);
  WriteLn(s.Pop);
  WriteLn(s.Pop);
  WriteLn(s.Pop);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["30", "20", "10"]);
}

#[test]
fn test_tstack_peek_top() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<String>;
begin
  s := TStack<String>.Create;
  s.Push('Bottom');
  s.Push('Top');
  WriteLn(s.Peek);
  WriteLn(s.Count);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["Top", "2"]);
}

#[test]
fn test_tqueue_clear_resets_queue() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<Integer>;
begin
  q := TQueue<Integer>.Create;
  q.Enqueue(1); q.Enqueue(2);
  q.Clear;
  WriteLn(q.Count);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tstack_clear_resets_stack() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<Integer>;
begin
  s := TStack<Integer>.Create;
  s.Push(1); s.Push(2);
  s.Clear;
  WriteLn(s.Count);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_tqueue_toarray_conversion() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<Integer>; arr: TArray<Integer>;
begin
  q := TQueue<Integer>.Create;
  q.Enqueue(5); q.Enqueue(15);
  arr := q.ToArray;
  WriteLn(Length(arr));
  WriteLn(arr[0]);
  WriteLn(arr[1]);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["2", "5", "15"]);
}

#[test]
fn test_tstack_toarray_conversion() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<Integer>; arr: TArray<Integer>;
begin
  s := TStack<Integer>.Create;
  s.Push(5); s.Push(15);
  arr := s.ToArray;
  WriteLn(Length(arr));
  s.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_tqueue_record_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TTask = record ID: Integer; Name: String; end;
var q: TQueue<TTask>; t1, t2: TTask;
begin
  q := TQueue<TTask>.Create;
  t1.ID := 1; t1.Name := 'TaskOne';
  q.Enqueue(t1);
  t2 := q.Dequeue;
  WriteLn(t2.ID);
  WriteLn(t2.Name);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["1", "TaskOne"]);
}

#[test]
fn test_tstack_record_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TPoint = record X, Y: Integer; end;
var s: TStack<TPoint>; pt, res: TPoint;
begin
  s := TStack<TPoint>.Create;
  pt.X := 100; pt.Y := 200;
  s.Push(pt);
  res := s.Pop;
  WriteLn(res.X);
  WriteLn(res.Y);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn test_tqueue_for_in_loop() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<String>; item: String;
begin
  q := TQueue<String>.Create;
  q.Enqueue('Q1'); q.Enqueue('Q2');
  for item in q do
    WriteLn(item);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["Q1", "Q2"]);
}

#[test]
fn test_tstack_for_in_loop() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<String>; item: String;
begin
  s := TStack<String>.Create;
  s.Push('S1'); s.Push('S2');
  for item in s do
    WriteLn(item);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["S2", "S1"]);
}

#[test]
fn test_tqueue_extract_front() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<Integer>; val: Integer;
begin
  q := TQueue<Integer>.Create;
  q.Enqueue(888);
  val := q.Extract;
  WriteLn(val);
  WriteLn(q.Count);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["888", "0"]);
}

#[test]
fn test_tstack_extract_top() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<Integer>; val: Integer;
begin
  s := TStack<Integer>.Create;
  s.Push(999);
  val := s.Extract;
  WriteLn(val);
  WriteLn(s.Count);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["999", "0"]);
}

#[test]
fn test_tqueue_trimexcess() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<Integer>;
begin
  q := TQueue<Integer>.Create;
  q.Enqueue(10); q.Enqueue(20);
  q.TrimExcess;
  WriteLn(q.Count);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_tstack_trimexcess() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<Integer>;
begin
  s := TStack<Integer>.Create;
  s.Push(10); s.Push(20);
  s.TrimExcess;
  WriteLn(s.Count);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_tqueue_interleaved_enqueue_dequeue() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var q: TQueue<Integer>;
begin
  q := TQueue<Integer>.Create;
  q.Enqueue(1);
  q.Enqueue(2);
  WriteLn(q.Dequeue);
  q.Enqueue(3);
  WriteLn(q.Dequeue);
  WriteLn(q.Dequeue);
  q.Free;
end.
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_tstack_interleaved_push_pop() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<Integer>;
begin
  s := TStack<Integer>.Create;
  s.Push(1);
  s.Push(2);
  WriteLn(s.Pop);
  s.Push(3);
  WriteLn(s.Pop);
  WriteLn(s.Pop);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["2", "3", "1"]);
}

#[test]
fn test_tqueue_enum_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
type TState = (sInit, sActive, sDone);
var q: TQueue<TState>;
begin
  q := TQueue<TState>.Create;
  q.Enqueue(sInit); q.Enqueue(sDone);
  WriteLn(Ord(q.Dequeue));
  WriteLn(Ord(q.Dequeue));
  q.Free;
end.
"#);
    assert_eq!(out, vec!["0", "2"]);
}

#[test]
fn test_tstack_real_elements() {
    let out = run_pascal(r#"
program Test;
uses Generics.Collections;
var s: TStack<Real>;
begin
  s := TStack<Real>.Create;
  s.Push(12.5); s.Push(25.0);
  WriteLn(s.Pop);
  WriteLn(s.Pop);
  s.Free;
end.
"#);
    assert_eq!(out, vec!["25", "12.5"]);
}
