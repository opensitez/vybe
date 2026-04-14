/// C# LINQ method syntax, lambda expressions, delegates, events,
/// Func<T>/Action<T>, higher-order functions, closures.

use super::helpers::run_csharp;

// ===================================================================
// LINQ METHOD CHAINS
// ===================================================================

#[test] fn linq_where_select_tolist() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
var result = nums.Where(x => x % 2 == 0).Select(x => x * 10).ToList();
foreach (var x in result) Console.WriteLine(x);
"#), &["20", "40", "60", "80"]);
}

#[test] fn linq_orderby_thenby() {
    assert_eq!(run_csharp(r#"
var names = new List<string> { "Charlie", "Alice", "Bob", "Alice" };
var sorted = names.OrderBy(n => n).ToList();
foreach (var n in sorted) Console.WriteLine(n);
"#), &["Alice", "Alice", "Bob", "Charlie"]);
}

#[test] fn linq_orderbydescending() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 3, 1, 4, 1, 5 };
var sorted = nums.OrderByDescending(x => x).ToList();
foreach (var x in sorted) Console.WriteLine(x);
"#), &["5", "4", "3", "1", "1"]);
}

#[test] fn linq_distinct() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 2, 3, 3, 3, 4 };
var distinct = nums.Distinct().ToList();
Console.WriteLine(distinct.Count);
"#), &["4"]);
}

#[test] fn linq_count_with_predicate() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 3, 4, 5, 6 };
Console.WriteLine(nums.Count(x => x > 3));
"#), &["3"]);
}

#[test] fn linq_sum() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 3, 4, 5 };
Console.WriteLine(nums.Sum());
"#), &["15"]);
}

#[test] fn linq_average() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 10, 20, 30 };
Console.WriteLine(nums.Average());
"#), &["20"]);
}

#[test] fn linq_min_max() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 5, 3, 8, 1, 9 };
Console.WriteLine(nums.Min());
Console.WriteLine(nums.Max());
"#), &["1", "9"]);
}

#[test] fn linq_first_last() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 10, 20, 30 };
Console.WriteLine(nums.First());
Console.WriteLine(nums.Last());
"#), &["10", "30"]);
}

#[test] fn linq_firstordefault_empty() {
    assert_eq!(run_csharp(r#"
var nums = new List<int>();
Console.WriteLine(nums.FirstOrDefault());
"#), &["0"]);
}

#[test] fn linq_any_all() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 2, 4, 6, 8 };
Console.WriteLine(nums.All(x => x % 2 == 0));
Console.WriteLine(nums.Any(x => x > 5));
Console.WriteLine(nums.Any(x => x > 10));
"#), &["True", "True", "False"]);
}

#[test] fn linq_skip_take() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 3, 4, 5, 6, 7, 8 };
var page = nums.Skip(2).Take(3).ToList();
foreach (var x in page) Console.WriteLine(x);
"#), &["3", "4", "5"]);
}

#[test] fn linq_selectmany() {
    assert_eq!(run_csharp(r#"
var lists = new List<List<int>> {
    new List<int> { 1, 2 },
    new List<int> { 3, 4 },
    new List<int> { 5 }
};
var flat = lists.SelectMany(l => l).ToList();
foreach (var x in flat) Console.WriteLine(x);
"#), &["1", "2", "3", "4", "5"]);
}

#[test] fn linq_groupby() {
    assert_eq!(run_csharp(r#"
var words = new List<string> { "apple", "ant", "banana", "avocado", "bat" };
var groups = words.GroupBy(w => w[0].ToString()).ToList();
foreach (var g in groups) {
    Console.WriteLine(g.Key + ": " + g.Count());
}
"#), &["a: 3", "b: 2"]);
}

#[test] fn linq_zip() {
    assert_eq!(run_csharp(r#"
var names = new List<string> { "Alice", "Bob", "Charlie" };
var ages = new List<int> { 30, 25, 35 };
var pairs = names.Zip(ages, (n, a) => n + "=" + a).ToList();
foreach (var p in pairs) Console.WriteLine(p);
"#), &["Alice=30", "Bob=25", "Charlie=35"]);
}

#[test] fn linq_aggregate() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 3, 4, 5 };
var product = nums.Aggregate(1, (acc, x) => acc * x);
Console.WriteLine(product);
"#), &["120"]);
}

#[test] fn linq_todictionary() {
    assert_eq!(run_csharp(r#"
var names = new List<string> { "Alice", "Bob" };
var dict = names.ToDictionary(n => n, n => n.Length);
Console.WriteLine(dict["Alice"]);
Console.WriteLine(dict["Bob"]);
"#), &["5", "3"]);
}

#[test] fn linq_contains() {
    assert_eq!(run_csharp(r#"
var nums = new List<int> { 1, 2, 3, 4 };
Console.WriteLine(nums.Contains(3));
Console.WriteLine(nums.Contains(9));
"#), &["True", "False"]);
}

// ===================================================================
// LAMBDA EXPRESSIONS
// ===================================================================

#[test] fn lambda_multiline() {
    assert_eq!(run_csharp(r#"
Func<int, int> factorial = null;
factorial = n => {
    if (n <= 1) return 1;
    return n * factorial(n - 1);
};
Console.WriteLine(factorial(5));
"#), &["120"]);
}

#[test] fn lambda_with_capture() {
    assert_eq!(run_csharp(r#"
int multiplier = 3;
Func<int, int> mul = x => x * multiplier;
Console.WriteLine(mul(10));
Console.WriteLine(mul(7));
"#), &["30", "21"]);
}

#[test] fn lambda_passed_to_method() {
    assert_eq!(run_csharp(r#"
class Processor {
    public int Apply(int value, Func<int, int> transform) {
        return transform(value);
    }
}
var p = new Processor();
Console.WriteLine(p.Apply(5, x => x * x));
Console.WriteLine(p.Apply(5, x => x + 10));
"#), &["25", "15"]);
}

// ===================================================================
// ACTION<T> / FUNC<T>
// ===================================================================

#[test] fn action_with_params() {
    assert_eq!(run_csharp(r#"
Action<string, int> describe = (name, age) => {
    Console.WriteLine(name + " is " + age);
};
describe("Alice", 30);
describe("Bob", 25);
"#), &["Alice is 30", "Bob is 25"]);
}

#[test] fn func_chain() {
    assert_eq!(run_csharp(r#"
Func<int, int> doubleIt = x => x * 2;
Func<int, int> addOne = x => x + 1;
Console.WriteLine(addOne(doubleIt(5)));
Console.WriteLine(doubleIt(addOne(5)));
"#), &["11", "12"]);
}

#[test] fn func_as_return_value() {
    assert_eq!(run_csharp(r#"
Func<int, Func<int, int>> makeAdder = x => y => x + y;
var add5 = makeAdder(5);
Console.WriteLine(add5(3));
Console.WriteLine(add5(10));
"#), &["8", "15"]);
}

// ===================================================================
// EVENTS AND DELEGATES
// ===================================================================

#[test] fn delegate_multicast() {
    assert_eq!(run_csharp(r#"
Action<string> logger = msg => Console.WriteLine("LOG: " + msg);
Action<string> printer = msg => Console.WriteLine("PRINT: " + msg);
Action<string> both = logger + printer;
both("hello");
"#), &["LOG: hello", "PRINT: hello"]);
}

#[test] fn event_pattern() {
    assert_eq!(run_csharp(r#"
class Button {
    public event Action OnClick;
    public void Click() { if (OnClick != null) OnClick(); }
}
var btn = new Button();
btn.OnClick += () => Console.WriteLine("clicked!");
btn.Click();
btn.Click();
"#), &["clicked!", "clicked!"]);
}

#[test] fn event_with_args() {
    assert_eq!(run_csharp(r#"
class Timer {
    public event Action<int> OnTick;
    public void Tick(int count) { if (OnTick != null) OnTick(count); }
}
var t = new Timer();
t.OnTick += n => Console.WriteLine("tick " + n);
t.Tick(1);
t.Tick(2);
"#), &["tick 1", "tick 2"]);
}

// ===================================================================
// PREDICATE<T>
// ===================================================================

#[test] fn predicate_usage() {
    assert_eq!(run_csharp(r#"
Predicate<int> isEven = x => x % 2 == 0;
Console.WriteLine(isEven(4));
Console.WriteLine(isEven(7));
"#), &["True", "False"]);
}
