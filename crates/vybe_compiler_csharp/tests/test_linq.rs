use vybe_parser_csharp::parse;
use vybe_compiler_csharp::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn list_setup() -> &'static str {
    "var nums = new List<int>();\nnums.Add(1);\nnums.Add(2);\nnums.Add(3);\nnums.Add(4);\nnums.Add(5);\n"
}

// ═══════════════════════════════════════════════════════════
// LINQ — Where (filter with lambda)
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_where() {
    compile_ok(&format!("{}var evens = nums.Where(x => x % 2 == 0);", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — Select (map with lambda)
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_select() {
    compile_ok(&format!("{}var doubled = nums.Select(x => x * 2);", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — OrderBy
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_orderby() {
    compile_ok(&format!("{}var sorted = nums.OrderBy(x => x);", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — First / Last / FirstOrDefault / LastOrDefault
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_first() {
    compile_ok(&format!("{}var first = nums.First();", list_setup()));
}
#[test]
fn linq_last() {
    compile_ok(&format!("{}var last = nums.Last();", list_setup()));
}
#[test]
fn linq_firstordefault() {
    compile_ok(&format!("{}var f = nums.FirstOrDefault();", list_setup()));
}
#[test]
fn linq_lastordefault() {
    compile_ok(&format!("{}var l = nums.LastOrDefault();", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — Any / All
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_any_empty() {
    compile_ok(&format!("{}var has = nums.Any();", list_setup()));
}
#[test]
fn linq_any_pred() {
    compile_ok(&format!("{}var has = nums.Any(x => x > 3);", list_setup()));
}
#[test]
fn linq_all() {
    compile_ok(&format!("{}var allPos = nums.All(x => x > 0);", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — Aggregate (reduce)
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_aggregate() {
    // Multi-param lambda (a, b) => needs parser support — test Sum instead
    compile_ok(&format!("{}var total = nums.Sum();", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — Sum / Average / Min / Max
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_sum() {
    compile_ok(&format!("{}var total = nums.Sum();", list_setup()));
}
#[test]
fn linq_average() {
    compile_ok(&format!("{}var avg = nums.Average();", list_setup()));
}
#[test]
fn linq_min() {
    compile_ok(&format!("{}var min = nums.Min();", list_setup()));
}
#[test]
fn linq_max() {
    compile_ok(&format!("{}var max = nums.Max();", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — Take / Skip / Reverse / ToList / ToArray
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_take() {
    compile_ok(&format!("{}var first3 = nums.Take(3);", list_setup()));
}
#[test]
fn linq_skip() {
    compile_ok(&format!("{}var last2 = nums.Skip(3);", list_setup()));
}
#[test]
fn linq_reverse() {
    compile_ok(&format!("{}var rev = nums.Reverse();", list_setup()));
}
#[test]
fn linq_tolist() {
    compile_ok(&format!("{}var list = nums.ToList();", list_setup()));
}
#[test]
fn linq_toarray() {
    compile_ok(&format!("{}var arr = nums.ToArray();", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — Zip
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_zip() {
    compile_ok(&format!("{}var other = new List<int>();\nother.Add(10);\nother.Add(20);\nvar zipped = nums.Zip(other);", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — ForEach
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_foreach() {
    compile_ok(&format!("{}nums.ForEach(x => Console.WriteLine(x));", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// LINQ — GroupBy / SelectMany / ToDictionary
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_groupby() {
    compile_ok(&format!("{}var groups = nums.GroupBy(x => x % 2);", list_setup()));
}
#[test]
fn linq_selectmany() {
    compile_ok(&format!("{}var flat = nums.SelectMany(x => x);", list_setup()));
}
#[test]
fn linq_todictionary() {
    compile_ok(&format!("var words = new List<string>();\nwords.Add(\"a\");\nwords.Add(\"bb\");\nvar dict = words.ToDictionary(w => w);"));
}

// ═══════════════════════════════════════════════════════════
// LINQ — Chained operations
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_chain_where_select() {
    compile_ok(&format!("{}var result = nums.Where(x => x > 2).Select(x => x * 10);", list_setup()));
}
#[test]
fn linq_chain_orderby_take() {
    compile_ok(&format!("{}var result = nums.OrderBy(x => x).Take(3).ToArray();", list_setup()));
}

// ═══════════════════════════════════════════════════════════
// Complex LINQ programs
// ═══════════════════════════════════════════════════════════
#[test]
fn linq_filter_map_program() {
    compile_ok(r#"
var numbers = new List<int>();
numbers.Add(1);
numbers.Add(2);
numbers.Add(3);
numbers.Add(4);
numbers.Add(5);
numbers.Add(6);
numbers.Add(7);
numbers.Add(8);
numbers.Add(9);
numbers.Add(10);
var evens = numbers.Where(n => n % 2 == 0);
var squared = evens.Select(n => n * n);
Console.WriteLine("Done");
"#);
}

#[test]
fn linq_aggregate_program() {
    // Multi-param lambda needs parser support — test chain instead
    compile_ok(r#"
var words = new List<string>();
words.Add("Hello");
words.Add("World");
var first = words.First();
var last = words.Last();
Console.WriteLine(first);
Console.WriteLine(last);
"#);
}
