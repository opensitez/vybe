use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value};

fn run_cs(source: &str) -> Vec<String> {
    let unit = vybe_parser_csharp::parse(source).unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybe_host::setup_namespaces(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_vm: &mut VM, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    let chunks = vybe_compiler_csharp::Compiler::new().compile(&unit)
        .unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));
    let result = output.borrow().clone();
    result
}

// ═══════════════════════════════════════════════════════════
// LINQ Where — filters with VM callback
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_where_even() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3); list.Add(4); list.Add(5);
var evens = list.Where(x => x % 2 == 0);
evens.ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["2", "4"]);
}

#[test]
fn linq_where_gt() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(10); list.Add(20); list.Add(30); list.Add(40); list.Add(50);
list.Where(x => x > 25).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["30", "40", "50"]);
}

// ═══════════════════════════════════════════════════════════
// LINQ Select — maps with VM callback
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_select_double() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
list.Select(x => x * 2).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["2", "4", "6"]);
}

#[test]
fn linq_select_strings() {
    let out = run_cs(r#"
var list = new List<string>();
list.Add("hello"); list.Add("world");
list.Select(s => s.ToUpper()).ForEach(s => Console.WriteLine(s));
"#);
    assert_eq!(out, ["HELLO", "WORLD"]);
}

// ═══════════════════════════════════════════════════════════
// LINQ Any — with predicate callback
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_any_pred() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
Console.WriteLine(list.Any(x => x > 10));
Console.WriteLine(list.Any(x => x == 2));
"#);
    assert_eq!(out, ["false", "true"]);
}

// ═══════════════════════════════════════════════════════════
// LINQ All — with predicate callback
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_all_pred() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
Console.WriteLine(list.All(x => x > 0));
Console.WriteLine(list.All(x => x > 2));
"#);
    assert_eq!(out, ["true", "false"]);
}

// ═══════════════════════════════════════════════════════════
// LINQ ForEach — callback on each element
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_foreach() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(10); list.Add(20); list.Add(30);
list.ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["10", "20", "30"]);
}

// ═══════════════════════════════════════════════════════════
// LINQ Aggregate — reduce with VM callback
// ═══════════════════════════════════════════════════════════
// Note: multi-param lambda (a, b) => not yet supported by parser

// ═══════════════════════════════════════════════════════════
// LINQ Chained: Where + Select + ForEach
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_chain_where_select() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3); list.Add(4); list.Add(5); list.Add(6);
list.Where(x => x % 2 == 0).Select(x => x * 10).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["20", "40", "60"]);
}

#[test]
fn linq_chain_where_foreach() {
    let out = run_cs(r#"
var nums = new List<int>();
nums.Add(1); nums.Add(2); nums.Add(3); nums.Add(4); nums.Add(5);
nums.Where(x => x > 3).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["4", "5"]);
}

// ═══════════════════════════════════════════════════════════
// LINQ OrderBy
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_orderby() {
    let out = run_cs(r#"
var list = new List<int>();
list.Add(5); list.Add(3); list.Add(1); list.Add(4); list.Add(2);
list.OrderBy(x => x).ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["1", "2", "3", "4", "5"]);
}

// ═══════════════════════════════════════════════════════════
// LINQ complex program
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_filter_map_program() {
    let out = run_cs(r#"
var numbers = new List<int>();
numbers.Add(1); numbers.Add(2); numbers.Add(3); numbers.Add(4);
numbers.Add(5); numbers.Add(6); numbers.Add(7); numbers.Add(8);
var result = numbers.Where(n => n % 2 == 0).Select(n => n * n);
result.ForEach(x => Console.WriteLine(x));
"#);
    assert_eq!(out, ["4", "16", "36", "64"]);
}
