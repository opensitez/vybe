use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Arrays — new[] syntax, multi-dim, iteration, methods
// ═══════════════════════════════════════════════════════════

#[test]
fn new_array_syntax() {
    let out = run_csharp(
        r#"
var arr = new[] { 10, 20, 30, 40, 50 };
Console.WriteLine(arr.Length);
Console.WriteLine(arr[2]);
"#,
    );
    assert_eq!(out, vec!["5", "30"]);
}

#[test]
fn array_foreach() {
    let out = run_csharp(
        r#"
var names = new[] { "Alice", "Bob", "Carol" };
foreach (var name in names) {
    Console.WriteLine(name);
}
"#,
    );
    assert_eq!(out, vec!["Alice", "Bob", "Carol"]);
}

#[test]
fn array_for_loop() {
    let out = run_csharp(
        r#"
var arr = new[] { 1, 2, 3, 4, 5 };
int sum = 0;
for (int i = 0; i < arr.Length; i++) {
    sum += arr[i];
}
Console.WriteLine(sum);
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn array_set_values() {
    let out = run_csharp(
        r#"
var arr = new int[3];
arr[0] = 10;
arr[1] = 20;
arr[2] = 30;
Console.WriteLine(arr[0] + arr[1] + arr[2]);
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn array_of_strings() {
    let out = run_csharp(
        r#"
var words = new[] { "hello", "world" };
Console.WriteLine(words[0] + " " + words[1]);
Console.WriteLine(words.Length);
"#,
    );
    assert_eq!(out, vec!["hello world", "2"]);
}

#[test]
fn array_passed_to_method() {
    let out = run_csharp(
        r#"
class Utils {
    public static int Sum(int[] arr) {
        int total = 0;
        foreach (var x in arr) total += x;
        return total;
    }
}
var nums = new[] { 1, 2, 3, 4 };
Console.WriteLine(Utils.Sum(nums));
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn array_of_objects() {
    let out = run_csharp(
        r#"
class Item {
    public string Name;
    public Item(string n) { Name = n; }
}
var items = new[] { new Item("a"), new Item("b"), new Item("c") };
foreach (var item in items) {
    Console.WriteLine(item.Name);
}
"#,
    );
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn string_join_array() {
    let out = run_csharp(
        r#"
var arr = new[] { "a", "b", "c" };
Console.WriteLine(string.Join(",", arr));
Console.WriteLine(string.Join(" - ", arr));
"#,
    );
    assert_eq!(out, vec!["a,b,c", "a - b - c"]);
}
