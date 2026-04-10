use super::helpers::run_csharp;

#[test]
fn linq_where_even() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3); list.Add(4); list.Add(5);
var evens = list.Where(x => x % 2 == 0);
evens.ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["2", "4"]);
}

#[test]
fn linq_where_gt() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(10); list.Add(20); list.Add(30); list.Add(40); list.Add(50);
list.Where(x => x > 25).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["30", "40", "50"]);
}

#[test]
fn linq_select_double() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
list.Select(x => x * 2).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["2", "4", "6"]);
}

#[test]
fn linq_select_strings() {
    let out = run_csharp(r#"
var list = new List<string>();
list.Add("hello"); list.Add("world");
list.Select(s => s.ToUpper()).ForEach(s => Console.WriteLine(s));
"#);
    assert_eq!(out, ["HELLO", "WORLD"]);
}

#[test]
fn linq_any_pred() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
Console.WriteLine(list.Any(x => x > 10));
Console.WriteLine(list.Any(x => x == 2));
"#);
    assert_eq!(out, ["false", "true"]);
}

#[test]
fn linq_all_pred() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
Console.WriteLine(list.All(x => x > 0));
Console.WriteLine(list.All(x => x > 2));
"#);
    assert_eq!(out, ["true", "false"]);
}

#[test]
fn linq_foreach_runtime() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(10); list.Add(20); list.Add(30);
list.ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["10", "20", "30"]);
}

#[test]
fn linq_chain_where_select_foreach() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3); list.Add(4); list.Add(5); list.Add(6);
list.Where(x => x % 2 == 0).Select(x => x * 10).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["20", "40", "60"]);
}

#[test]
fn linq_chain_where_foreach() {
    let out = run_csharp(r#"
var nums = new List<int>();
nums.Add(1); nums.Add(2); nums.Add(3); nums.Add(4); nums.Add(5);
nums.Where(x => x > 3).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["4", "5"]);
}

#[test]
fn linq_orderby() {
    let out = run_csharp(r#"
var list = new List<int>();
list.Add(5); list.Add(3); list.Add(1); list.Add(4); list.Add(2);
list.OrderBy(x => x).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["1", "2", "3", "4", "5"]);
}

#[test]
fn linq_filter_map_program() {
    let out = run_csharp(r#"
var numbers = new List<int>();
numbers.Add(1); numbers.Add(2); numbers.Add(3); numbers.Add(4);
numbers.Add(5); numbers.Add(6); numbers.Add(7); numbers.Add(8);
var result = numbers.Where(n => n % 2 == 0).Select(n => n * n);
result.ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["4", "16", "36", "64"]);
}
