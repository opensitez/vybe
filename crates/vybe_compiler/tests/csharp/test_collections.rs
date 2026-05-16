use super::helpers::{run_csharp, run_csharp_one};

// ── List basic operations ───────────────────────────────────────────────────

#[test]
fn list_add_and_iterate() {
    let out = run_csharp(r#"
        var list = new List<string>();
        list.Add("a");
        list.Add("b");
        list.Add("c");
        foreach (var item in list) {
            Console.WriteLine(item);
        }
    "#);
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn list_add_numbers() {
    let out = run_csharp(r#"
        var list = new List<int>();
        list.Add(10);
        list.Add(20);
        list.Add(30);
        var sum = 0;
        foreach (var x in list) { sum = sum + x; }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["60"]);
}

#[test]
fn list_index_access() {
    let out = run_csharp(r#"
        var list = new List<int>();
        list.Add(10);
        list.Add(20);
        list.Add(30);
        Console.WriteLine(list[0]);
        Console.WriteLine(list[2]);
    "#);
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn list_sort() {
    let out = run_csharp(r#"
        var list = new List<int>();
        list.Add(3);
        list.Add(1);
        list.Add(2);
        list.Sort();
        foreach (var x in list) { Console.WriteLine(x); }
    "#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn list_reverse() {
    let out = run_csharp(r#"
        var list = new List<int>();
        list.Add(1);
        list.Add(2);
        list.Add(3);
        list.Reverse();
        foreach (var x in list) { Console.WriteLine(x); }
    "#);
    assert_eq!(out, vec!["3", "2", "1"]);
}

// ── Queue ───────────────────────────────────────────────────────────────────

#[test]
fn queue_enqueue_dequeue() {
    let out = run_csharp(r#"
        var q = new Queue<string>();
        q.Enqueue("first");
        q.Enqueue("second");
        Console.WriteLine(q.Dequeue());
        Console.WriteLine(q.Dequeue());
    "#);
    assert_eq!(out, vec!["first", "second"]);
}

// ── Stack ───────────────────────────────────────────────────────────────────

#[test]
fn stack_push_pop() {
    let out = run_csharp(r#"
        var s = new Stack<int>();
        s.Push(1);
        s.Push(2);
        s.Push(3);
        Console.WriteLine(s.Pop());
        Console.WriteLine(s.Pop());
    "#);
    assert_eq!(out, vec!["3", "2"]);
}

// ── Dictionary ──────────────────────────────────────────────────────────────

#[test]
fn dictionary_basic() {
    let out = run_csharp(r#"
        var dict = new Dictionary<string, int>();
        dict.Add("x", 10);
        dict.Add("y", 20);
        Console.WriteLine(dict["x"]);
        Console.WriteLine(dict["y"]);
    "#);
    assert_eq!(out, vec!["10", "20"]);
}
