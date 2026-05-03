/// C# collections: List<T>, Dictionary<K,V>, HashSet<T>, Queue<T>, Stack<T>,
/// SortedList, LinkedList, array patterns, collection initializers.

use super::helpers::run_csharp;

// ===================================================================
// LIST<T> OPERATIONS
// ===================================================================

#[test] fn list_addrange() {
    assert_eq!(run_csharp(r#"
var list = new List<int> { 1, 2, 3 };
list.AddRange(new int[] { 4, 5 });
Console.WriteLine(list.Count);
foreach (var x in list) Console.WriteLine(x);
"#), &["5", "1", "2", "3", "4", "5"]);
}

#[test] fn list_insert_at() {
    assert_eq!(run_csharp(r#"
var list = new List<string> { "a", "c", "d" };
list.Insert(1, "b");
foreach (var s in list) Console.WriteLine(s);
"#), &["a", "b", "c", "d"]);
}

#[test] fn list_removeat() {
    assert_eq!(run_csharp(r#"
var list = new List<int> { 10, 20, 30, 40 };
list.RemoveAt(1);
foreach (var x in list) Console.WriteLine(x);
"#), &["10", "30", "40"]);
}

#[test] fn list_find() {
    assert_eq!(run_csharp(r#"
var list = new List<int> { 1, 2, 3, 4, 5 };
var found = list.Find(x => x > 3);
Console.WriteLine(found);
"#), &["4"]);
}

#[test] fn list_findall() {
    assert_eq!(run_csharp(r#"
var list = new List<int> { 1, 2, 3, 4, 5, 6 };
var evens = list.FindAll(x => x % 2 == 0);
foreach (var x in evens) Console.WriteLine(x);
"#), &["2", "4", "6"]);
}

#[test] fn list_exists() {
    assert_eq!(run_csharp(r#"
var list = new List<string> { "apple", "banana", "cherry" };
Console.WriteLine(list.Exists(s => s == "banana"));
Console.WriteLine(list.Exists(s => s == "grape"));
"#), &["True", "False"]);
}

#[test] fn list_trueforall() {
    assert_eq!(run_csharp(r#"
var list = new List<int> { 2, 4, 6, 8 };
Console.WriteLine(list.TrueForAll(x => x % 2 == 0));
list.Add(3);
Console.WriteLine(list.TrueForAll(x => x % 2 == 0));
"#), &["True", "False"]);
}

#[test] fn list_convertall() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 3 };
var strings = nums.ConvertAll(x => x.ToString());
foreach (var s in strings) Console.WriteLine(s);
"#), &["1", "2", "3"]);
}

#[test] fn list_toarray() {
    assert_eq!(run_csharp(r#"
var list = new List<int> { 10, 20, 30 };
int[] arr = list.ToArray();
Console.WriteLine(arr.Length);
Console.WriteLine(arr[1]);
"#), &["3", "20"]);
}

// ===================================================================
// DICTIONARY<K,V>
// ===================================================================

#[test] fn dict_add_and_access() {
    assert_eq!(run_csharp(r#"
var dict = new Dictionary<string, int>();
dict.Add("one", 1);
dict.Add("two", 2);
dict.Add("three", 3);
Console.WriteLine(dict["two"]);
Console.WriteLine(dict.Count);
"#), &["2", "3"]);
}

#[test] fn dict_containskey() {
    assert_eq!(run_csharp(r#"
var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 } };
Console.WriteLine(dict.ContainsKey("a"));
Console.WriteLine(dict.ContainsKey("c"));
"#), &["True", "False"]);
}

#[test] fn dict_containsvalue() {
    assert_eq!(run_csharp(r#"
var dict = new Dictionary<string, int> { { "x", 10 }, { "y", 20 } };
Console.WriteLine(dict.ContainsValue(10));
Console.WriteLine(dict.ContainsValue(30));
"#), &["True", "False"]);
}

#[test] fn dict_trygetvalue() {
    assert_eq!(run_csharp(r#"
var dict = new Dictionary<string, int> { { "age", 30 } };
int value;
if (dict.TryGetValue("age", out value)) {
    Console.WriteLine(value);
}
if (!dict.TryGetValue("name", out value)) {
    Console.WriteLine("not found");
}
"#), &["30", "not found"]);
}

#[test] fn dict_iterate_keys_values() {
    assert_eq!(run_csharp(r#"
var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 } };
foreach (var key in dict.Keys) Console.WriteLine(key);
"#), &["a", "b"]);
}

#[test] fn dict_remove() {
    assert_eq!(run_csharp(r#"
var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 }, { "c", 3 } };
dict.Remove("b");
Console.WriteLine(dict.Count);
Console.WriteLine(dict.ContainsKey("b"));
"#), &["2", "False"]);
}

// ===================================================================
// HASHSET<T>
// ===================================================================

#[test] fn hashset_basic() {
    assert_eq!(run_csharp(r#"
var set = new HashSet<int> { 1, 2, 3, 2, 1 };
Console.WriteLine(set.Count);
Console.WriteLine(set.Contains(2));
Console.WriteLine(set.Contains(5));
"#), &["3", "True", "False"]);
}

#[test] fn hashset_add_remove() {
    assert_eq!(run_csharp(r#"
var set = new HashSet<string>();
set.Add("apple");
set.Add("banana");
set.Add("apple");
Console.WriteLine(set.Count);
set.Remove("apple");
Console.WriteLine(set.Count);
Console.WriteLine(set.Contains("apple"));
"#), &["2", "1", "False"]);
}

#[test] fn hashset_union() {
    assert_eq!(run_csharp(r#"
var a = new HashSet<int> { 1, 2, 3 };
var b = new HashSet<int> { 3, 4, 5 };
a.UnionWith(b);
Console.WriteLine(a.Count);
"#), &["5"]);
}

#[test] fn hashset_intersect() {
    assert_eq!(run_csharp(r#"
var a = new HashSet<int> { 1, 2, 3, 4 };
var b = new HashSet<int> { 2, 4, 6 };
a.IntersectWith(b);
Console.WriteLine(a.Count);
Console.WriteLine(a.Contains(2));
Console.WriteLine(a.Contains(4));
"#), &["2", "True", "True"]);
}

// ===================================================================
// QUEUE<T>
// ===================================================================

#[test] fn queue_enqueue_dequeue() {
    assert_eq!(run_csharp(r#"
var q = new Queue<string>();
q.Enqueue("first");
q.Enqueue("second");
q.Enqueue("third");
Console.WriteLine(q.Dequeue());
Console.WriteLine(q.Dequeue());
Console.WriteLine(q.Count);
"#), &["first", "second", "1"]);
}

#[test] fn queue_peek() {
    assert_eq!(run_csharp(r#"
var q = new Queue<int>();
q.Enqueue(10);
q.Enqueue(20);
Console.WriteLine(q.Peek());
Console.WriteLine(q.Count);
"#), &["10", "2"]);
}

// ===================================================================
// STACK<T>
// ===================================================================

#[test] fn stack_push_pop_peek() {
    assert_eq!(run_csharp(r#"
var s = new Stack<int>();
s.Push(1);
s.Push(2);
s.Push(3);
Console.WriteLine(s.Peek());
Console.WriteLine(s.Pop());
Console.WriteLine(s.Pop());
Console.WriteLine(s.Count);
"#), &["3", "3", "2", "1"]);
}

// ===================================================================
// ARRAY PATTERNS
// ===================================================================

#[test] fn array_initialization_syntax() {
    assert_eq!(run_csharp(r#"
int[] a = { 1, 2, 3, 4, 5 };
Console.WriteLine(a.Length);
Console.WriteLine(a[0] + a[4]);
"#), &["5", "6"]);
}

#[test] fn array_sort() {
    assert_eq!(run_csharp(r#"
int[] arr = { 5, 3, 8, 1, 2 };
Array.Sort(arr);
foreach (var x in arr) Console.WriteLine(x);
"#), &["1", "2", "3", "5", "8"]);
}

#[test] fn array_reverse() {
    assert_eq!(run_csharp(r#"
int[] arr = { 1, 2, 3, 4, 5 };
Array.Reverse(arr);
foreach (var x in arr) Console.WriteLine(x);
"#), &["5", "4", "3", "2", "1"]);
}

#[test] fn array_indexof() {
    assert_eq!(run_csharp(r#"
string[] arr = { "a", "b", "c", "d" };
Console.WriteLine(Array.IndexOf(arr, "c"));
Console.WriteLine(Array.IndexOf(arr, "z"));
"#), &["2", "-1"]);
}

#[test] fn array_exists() {
    assert_eq!(run_csharp(r#"
int[] arr = { 1, 2, 3, 4, 5 };
Console.WriteLine(Array.Exists(arr, x => x > 4));
Console.WriteLine(Array.Exists(arr, x => x > 10));
"#), &["True", "False"]);
}

#[test] fn array_find() {
    assert_eq!(run_csharp(r#"
int[] arr = { 10, 20, 30, 40 };
Console.WriteLine(Array.Find(arr, x => x > 15));
"#), &["20"]);
}

#[test] fn multidimensional_array() {
    assert_eq!(run_csharp(r#"
int[,] matrix = { { 1, 2 }, { 3, 4 }, { 5, 6 } };
Console.WriteLine(matrix[0, 0]);
Console.WriteLine(matrix[1, 1]);
Console.WriteLine(matrix[2, 0]);
"#), &["1", "4", "5"]);
}

#[test] fn jagged_array() {
    assert_eq!(run_csharp(r#"
int[][] jagged = new int[3][];
jagged[0] = new int[] { 1, 2 };
jagged[1] = new int[] { 3, 4, 5 };
jagged[2] = new int[] { 6 };
Console.WriteLine(jagged[0].Length);
Console.WriteLine(jagged[1].Length);
Console.WriteLine(jagged[1][2]);
"#), &["2", "3", "5"]);
}

// ===================================================================
// COLLECTION INITIALIZERS
// ===================================================================

#[test] fn collection_initializer_list() {
    assert_eq!(run_csharp(r#"
var names = new List<string> { "Alice", "Bob", "Charlie" };
Console.WriteLine(names.Count);
Console.WriteLine(names[1]);
"#), &["3", "Bob"]);
}

#[test] fn collection_initializer_dict() {
    assert_eq!(run_csharp(r#"
var ages = new Dictionary<string, int> {
    { "Alice", 30 },
    { "Bob", 25 }
};
Console.WriteLine(ages["Alice"]);
Console.WriteLine(ages.Count);
"#), &["30", "2"]);
}

