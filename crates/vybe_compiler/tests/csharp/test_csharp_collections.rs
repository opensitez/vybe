use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Collections — List, Dictionary, Queue, Stack, arrays
// ═══════════════════════════════════════════════════════════

#[test]
fn list_add_and_count() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var list = new List<int>();
list.Add(10);
list.Add(20);
list.Add(30);
Console.WriteLine(list.Count);
Console.WriteLine(list[1]);
"#);
    assert_eq!(out, vec!["3", "20"]);
}

#[test]
fn list_contains_remove() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var list = new List<string>();
list.Add("apple");
list.Add("banana");
list.Add("cherry");
Console.WriteLine(list.Contains("banana"));
list.Remove("banana");
Console.WriteLine(list.Count);
Console.WriteLine(list.Contains("banana"));
"#);
    assert_eq!(out, vec!["True", "2", "False"]);
}

#[test]
fn list_foreach() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var list = new List<int>();
list.Add(1);
list.Add(2);
list.Add(3);
int sum = 0;
foreach (var x in list) {
    sum += x;
}
Console.WriteLine(sum);
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn dictionary_basic() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var dict = new Dictionary<string, int>();
dict.Add("x", 10);
dict.Add("y", 20);
Console.WriteLine(dict["x"]);
Console.WriteLine(dict.ContainsKey("y"));
Console.WriteLine(dict.ContainsKey("z"));
Console.WriteLine(dict.Count);
"#);
    assert_eq!(out, vec!["10", "True", "False", "2"]);
}

#[test]
fn dictionary_remove() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var dict = new Dictionary<string, int>();
dict.Add("a", 1);
dict.Add("b", 2);
dict.Remove("a");
Console.WriteLine(dict.Count);
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn queue_operations() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
q.Enqueue("third");
Console.WriteLine(q.Count);
Console.WriteLine(q.Dequeue());
Console.WriteLine(q.Peek());
"#);
    assert_eq!(out, vec!["3", "first", "second"]);
}

#[test]
fn stack_operations() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var s = new Stack<int>();
s.Push(1);
s.Push(2);
s.Push(3);
Console.WriteLine(s.Count);
Console.WriteLine(s.Pop());
Console.WriteLine(s.Peek());
"#);
    assert_eq!(out, vec!["3", "3", "2"]);
}

#[test]
fn array_creation() {
    let out = run_csharp(r#"
int[] arr = {5, 10, 15, 20, 25};
Console.WriteLine(arr.Length);
Console.WriteLine(arr[2]);
"#);
    assert_eq!(out, vec!["5", "15"]);
}

#[test]
fn array_foreach() {
    let out = run_csharp(r#"
string[] names = {"Alice", "Bob", "Carol"};
foreach (var name in names) {
    Console.WriteLine(name);
}
"#);
    assert_eq!(out, vec!["Alice", "Bob", "Carol"]);
}

#[test]
fn list_sort_and_reverse() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var list = new List<int>();
list.Add(3);
list.Add(1);
list.Add(4);
list.Add(1);
list.Add(5);
list.Sort();
Console.WriteLine(list[0]);
Console.WriteLine(list[4]);
list.Reverse();
Console.WriteLine(list[0]);
"#);
    assert_eq!(out, vec!["1", "5", "5"]);
}

#[test]
fn list_indexof() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var list = new List<string>();
list.Add("a");
list.Add("b");
list.Add("c");
Console.WriteLine(list.IndexOf("b"));
Console.WriteLine(list.IndexOf("z"));
"#);
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn list_clear() {
    let out = run_csharp(r#"
using System.Collections.Generic;
var list = new List<int>();
list.Add(1);
list.Add(2);
Console.WriteLine(list.Count);
list.Clear();
Console.WriteLine(list.Count);
"#);
    assert_eq!(out, vec!["2", "0"]);
}
