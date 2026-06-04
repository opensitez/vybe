/// C# modern features: records, pattern matching (is, switch expressions,
/// property patterns), tuples, deconstruction, null-coalescing,
/// init-only setters, expression-bodied members, top-level statements.
use super::helpers::run_csharp;

// ===================================================================
// RECORDS
// ===================================================================

#[test]
fn record_basic() {
    assert_eq!(
        run_csharp(
            r#"
record Person(string Name, int Age);
var p = new Person("Alice", 30);
Console.WriteLine(p.Name);
Console.WriteLine(p.Age);
"#
        ),
        &["Alice", "30"]
    );
}

#[test]
fn record_tostring() {
    assert_eq!(
        run_csharp(
            r#"
record Point(int X, int Y);
var p = new Point(3, 4);
Console.WriteLine(p);
"#
        ),
        &["Point { X = 3, Y = 4 }"]
    );
}

#[test]
fn record_equality() {
    assert_eq!(
        run_csharp(
            r#"
record Color(int R, int G, int B);
var c1 = new Color(255, 0, 0);
var c2 = new Color(255, 0, 0);
var c3 = new Color(0, 255, 0);
Console.WriteLine(c1 == c2);
Console.WriteLine(c1 == c3);
"#
        ),
        &["True", "False"]
    );
}

#[test]
fn record_with_expression() {
    assert_eq!(
        run_csharp(
            r#"
record Person(string Name, int Age);
var p1 = new Person("Alice", 30);
var p2 = p1 with { Age = 31 };
Console.WriteLine(p1);
Console.WriteLine(p2);
"#
        ),
        &[
            "Person { Name = Alice, Age = 30 }",
            "Person { Name = Alice, Age = 31 }"
        ]
    );
}

// ===================================================================
// PATTERN MATCHING — IS PATTERN
// ===================================================================

#[test]
fn is_type_pattern() {
    assert_eq!(
        run_csharp(
            r#"
object obj = "hello";
if (obj is string s) {
    Console.WriteLine("string: " + s.ToUpper());
}
"#
        ),
        &["string: HELLO"]
    );
}

#[test]
fn is_constant_pattern() {
    assert_eq!(
        run_csharp(
            r#"
object obj = null;
Console.WriteLine(obj is null);
obj = 42;
Console.WriteLine(obj is 42);
Console.WriteLine(obj is 43);
"#
        ),
        &["True", "True", "False"]
    );
}

#[test]
fn is_not_pattern() {
    assert_eq!(
        run_csharp(
            r#"
object obj = "test";
if (obj is not null) {
    Console.WriteLine("not null");
}
if (obj is not int) {
    Console.WriteLine("not int");
}
"#
        ),
        &["not null", "not int"]
    );
}

// ===================================================================
// SWITCH EXPRESSION
// ===================================================================

#[test]
fn switch_expression_basic() {
    assert_eq!(
        run_csharp(
            r#"
int day = 3;
string name = day switch {
    1 => "Monday",
    2 => "Tuesday",
    3 => "Wednesday",
    4 => "Thursday",
    5 => "Friday",
    _ => "Weekend"
};
Console.WriteLine(name);
"#
        ),
        &["Wednesday"]
    );
}

#[test]
fn switch_expression_type_pattern() {
    assert_eq!(
        run_csharp(
            r#"
object obj = 42;
string result = obj switch {
    int i => "int: " + i,
    string s => "string: " + s,
    _ => "unknown"
};
Console.WriteLine(result);
"#
        ),
        &["int: 42"]
    );
}

#[test]
fn switch_expression_when_guard() {
    assert_eq!(
        run_csharp(
            r#"
int score = 85;
string grade = score switch {
    >= 90 => "A",
    >= 80 => "B",
    >= 70 => "C",
    >= 60 => "D",
    _ => "F"
};
Console.WriteLine(grade);
"#
        ),
        &["B"]
    );
}

// ===================================================================
// TUPLES
// ===================================================================

#[test]
fn tuple_basic() {
    assert_eq!(
        run_csharp(
            r#"
var t = (1, "hello", true);
Console.WriteLine(t.Item1);
Console.WriteLine(t.Item2);
Console.WriteLine(t.Item3);
"#
        ),
        &["1", "hello", "True"]
    );
}

#[test]
fn tuple_named() {
    assert_eq!(
        run_csharp(
            r#"
var p = (Name: "Alice", Age: 30);
Console.WriteLine(p.Name);
Console.WriteLine(p.Age);
"#
        ),
        &["Alice", "30"]
    );
}

#[test]
fn tuple_return_from_method() {
    assert_eq!(
        run_csharp(
            r#"
class MathOps {
    public static (int Min, int Max) MinMax(int[] arr) {
        int min = arr[0], max = arr[0];
        foreach (var x in arr) {
            if (x < min) min = x;
            if (x > max) max = x;
        }
        return (min, max);
    }
}
var result = MathOps.MinMax(new int[] { 3, 1, 4, 1, 5, 9 });
Console.WriteLine(result.Min);
Console.WriteLine(result.Max);
"#
        ),
        &["1", "9"]
    );
}

#[test]
fn tuple_deconstruction() {
    assert_eq!(
        run_csharp(
            r#"
var (name, age) = ("Bob", 25);
Console.WriteLine(name);
Console.WriteLine(age);
"#
        ),
        &["Bob", "25"]
    );
}

#[test]
fn tuple_equality() {
    assert_eq!(
        run_csharp(
            r#"
var t1 = (1, 2);
var t2 = (1, 2);
var t3 = (1, 3);
Console.WriteLine(t1 == t2);
Console.WriteLine(t1 == t3);
"#
        ),
        &["True", "False"]
    );
}

// ===================================================================
// NULL-COALESCING AND NULL-CONDITIONAL
// ===================================================================

#[test]
fn null_coalescing_operator() {
    assert_eq!(
        run_csharp(
            r#"
string s = null;
Console.WriteLine(s ?? "default");
s = "hello";
Console.WriteLine(s ?? "default");
"#
        ),
        &["default", "hello"]
    );
}

#[test]
fn null_coalescing_assignment() {
    assert_eq!(
        run_csharp(
            r#"
string s = null;
s ??= "assigned";
Console.WriteLine(s);
s ??= "not again";
Console.WriteLine(s);
"#
        ),
        &["assigned", "assigned"]
    );
}

#[test]
fn null_conditional_operator() {
    assert_eq!(
        run_csharp(
            r#"
string s = null;
Console.WriteLine(s?.ToUpper() ?? "null");
s = "hello";
Console.WriteLine(s?.ToUpper() ?? "null");
"#
        ),
        &["null", "HELLO"]
    );
}

#[test]
fn null_conditional_chain() {
    assert_eq!(
        run_csharp(
            r#"
class Inner { public string Value = "found"; }
class Outer { public Inner Child; }
var o = new Outer();
Console.WriteLine(o.Child?.Value ?? "missing");
o.Child = new Inner();
Console.WriteLine(o.Child?.Value ?? "missing");
"#
        ),
        &["missing", "found"]
    );
}

// ===================================================================
// TERNARY AND CONDITIONAL PATTERNS
// ===================================================================

#[test]
fn ternary_operator() {
    assert_eq!(
        run_csharp(
            r#"
int x = 10;
string result = x > 5 ? "big" : "small";
Console.WriteLine(result);
"#
        ),
        &["big"]
    );
}

#[test]
fn nested_ternary() {
    assert_eq!(
        run_csharp(
            r#"
int x = 50;
string cat = x < 0 ? "negative" : x == 0 ? "zero" : "positive";
Console.WriteLine(cat);
"#
        ),
        &["positive"]
    );
}

// ===================================================================
// EXPRESSION-BODIED MEMBERS
// ===================================================================

#[test]
fn expression_bodied_method_and_property() {
    assert_eq!(
        run_csharp(
            r#"
class Circle {
    public double Radius { get; }
    public Circle(double r) => Radius = r;
    public double Area => 3.14 * Radius * Radius;
    public double Circumference() => 2 * 3.14 * Radius;
}
var c = new Circle(5);
Console.WriteLine(c.Area);
Console.WriteLine(c.Circumference());
"#
        ),
        &["78.5", "31.4"]
    );
}

// ===================================================================
// RANGES AND INDICES
// ===================================================================

#[test]
fn range_operator_basic() {
    assert_eq!(
        run_csharp(
            r#"
int[] nums = { 1, 2, 3, 4, 5 };
int[] slice = nums[1..4];
foreach (var n in slice) Console.WriteLine(n);
"#
        ),
        &["2", "3", "4"]
    );
}

#[test]
fn index_from_end() {
    assert_eq!(
        run_csharp(
            r#"
int[] nums = { 10, 20, 30, 40, 50 };
Console.WriteLine(nums[^1]);
Console.WriteLine(nums[^2]);
"#
        ),
        &["50", "40"]
    );
}

#[test]
fn range_from_end() {
    assert_eq!(
        run_csharp(
            r#"
int[] nums = { 1, 2, 3, 4, 5 };
int[] last3 = nums[^3..];
foreach (var n in last3) Console.WriteLine(n);
"#
        ),
        &["3", "4", "5"]
    );
}

// ===================================================================
// VAR AND TYPE INFERENCE
// ===================================================================

#[test]
fn var_inference() {
    assert_eq!(
        run_csharp(
            r#"
var x = 42;
var s = "hello";
var list = new List<int> { 1, 2, 3 };
Console.WriteLine(x.GetType().Name);
Console.WriteLine(s.GetType().Name);
Console.WriteLine(list.Count);
"#
        ),
        &["Int32", "String", "3"]
    );
}

// ===================================================================
// TYPEOF / GETTYPE / NAMEOF
// ===================================================================

#[test]
fn typeof_gettype() {
    assert_eq!(
        run_csharp(
            r#"
Console.WriteLine(typeof(int).Name);
Console.WriteLine(typeof(string).Name);
Console.WriteLine(42.GetType().Name);
"#
        ),
        &["Int32", "String", "Int32"]
    );
}

#[test]
fn nameof_operator() {
    assert_eq!(
        run_csharp(
            r#"
int myVariable = 10;
Console.WriteLine(nameof(myVariable));
Console.WriteLine(nameof(Console));
"#
        ),
        &["myVariable", "Console"]
    );
}
