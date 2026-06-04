use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Generics — generic classes, methods, collections
// ═══════════════════════════════════════════════════════════

#[test]
fn generic_class() {
    let out = run_csharp(
        r#"
class Box<T> {
    public T Value;
    public Box(T val) { Value = val; }
}
var intBox = new Box<int>(42);
var strBox = new Box<string>("hello");
Console.WriteLine(intBox.Value);
Console.WriteLine(strBox.Value);
"#,
    );
    assert_eq!(out, vec!["42", "hello"]);
}

#[test]
fn generic_method() {
    let out = run_csharp(
        r#"
class Utils {
    public static T Identity<T>(T value) { return value; }
}
Console.WriteLine(Utils.Identity<int>(42));
Console.WriteLine(Utils.Identity<string>("hello"));
"#,
    );
    assert_eq!(out, vec!["42", "hello"]);
}

#[test]
fn generic_pair() {
    let out = run_csharp(
        r#"
class Pair<T1, T2> {
    public T1 First;
    public T2 Second;
    public Pair(T1 a, T2 b) { First = a; Second = b; }
}
var p = new Pair<string, int>("age", 30);
Console.WriteLine(p.First);
Console.WriteLine(p.Second);
"#,
    );
    assert_eq!(out, vec!["age", "30"]);
}

#[test]
fn generic_list_usage() {
    let out = run_csharp(
        r#"
var list = new List<int>();
list.Add(10);
list.Add(20);
list.Add(30);
Console.WriteLine(list.Count);
Console.WriteLine(list[1]);
"#,
    );
    assert_eq!(out, vec!["3", "20"]);
}

#[test]
fn generic_dictionary_usage() {
    let out = run_csharp(
        r#"
var dict = new Dictionary<string, int>();
dict.Add("x", 10);
dict.Add("y", 20);
Console.WriteLine(dict["x"]);
Console.WriteLine(dict.Count);
"#,
    );
    assert_eq!(out, vec!["10", "2"]);
}

#[test]
fn generic_stack_queue() {
    let out = run_csharp(
        r#"
var stack = new Stack<string>();
stack.Push("first");
stack.Push("second");
Console.WriteLine(stack.Pop());
var queue = new Queue<int>();
queue.Enqueue(1);
queue.Enqueue(2);
Console.WriteLine(queue.Dequeue());
"#,
    );
    assert_eq!(out, vec!["second", "1"]);
}
