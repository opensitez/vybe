/// C# common patterns: IDisposable, params arrays, named/optional
/// parameters, ref/out, static methods, enums, const/readonly,
/// array/list algorithms, math operations, type casting, for/foreach/while patterns.

use super::helpers::run_csharp;

// ===================================================================
// PARAMS ARRAYS
// ===================================================================

#[test] fn params_array_basic() {
    assert_eq!(run_csharp(r#"
class Logger {
    public static void Log(params string[] messages) {
        foreach (var m in messages) Console.WriteLine(m);
    }
}
Logger.Log("one", "two", "three");
"#), &["one", "two", "three"]);
}

#[test] fn params_with_normal_params() {
    assert_eq!(run_csharp(r#"
class Fmt {
    public static string Build(string prefix, params int[] nums) {
        return prefix + ": " + string.Join(",", nums);
    }
}
Console.WriteLine(Fmt.Build("nums", 1, 2, 3));
"#), &["nums: 1,2,3"]);
}

// ===================================================================
// NAMED AND OPTIONAL PARAMETERS
// ===================================================================

#[test] fn optional_params() {
    assert_eq!(run_csharp(r#"
class Greeter {
    public static string Hello(string name, string greeting = "Hello") {
        return greeting + ", " + name + "!";
    }
}
Console.WriteLine(Greeter.Hello("Alice"));
Console.WriteLine(Greeter.Hello("Bob", "Hi"));
"#), &["Hello, Alice!", "Hi, Bob!"]);
}

#[test] fn named_params() {
    assert_eq!(run_csharp(r#"
class Rect {
    public static int Area(int width, int height) { return width * height; }
}
Console.WriteLine(Rect.Area(width: 5, height: 3));
Console.WriteLine(Rect.Area(height: 10, width: 2));
"#), &["15", "20"]);
}

// ===================================================================
// REF AND OUT
// ===================================================================

#[test] fn ref_parameter() {
    assert_eq!(run_csharp(r#"
class Ops {
    public static void Double(ref int x) { x *= 2; }
}
int val = 5;
Ops.Double(ref val);
Console.WriteLine(val);
Ops.Double(ref val);
Console.WriteLine(val);
"#), &["10", "20"]);
}

#[test] fn out_parameter() {
    assert_eq!(run_csharp(r#"
class Parser {
    public static bool TryParse(string s, out int result) {
        if (s == "42") { result = 42; return true; }
        result = 0;
        return false;
    }
}
int val;
Console.WriteLine(Parser.TryParse("42", out val));
Console.WriteLine(val);
Console.WriteLine(Parser.TryParse("bad", out val));
Console.WriteLine(val);
"#), &["True", "42", "False", "0"]);
}

#[test] fn out_var_declaration() {
    assert_eq!(run_csharp(r#"
if (int.TryParse("123", out var result)) {
    Console.WriteLine(result);
}
"#), &["123"]);
}

// ===================================================================
// ENUMS
// ===================================================================

#[test] fn enum_basic() {
    assert_eq!(run_csharp(r#"
enum Color { Red, Green, Blue }
Color c = Color.Green;
Console.WriteLine(c);
Console.WriteLine((int)c);
"#), &["Green", "1"]);
}

#[test] fn enum_with_values() {
    assert_eq!(run_csharp(r#"
enum Status { Active = 1, Inactive = 0, Pending = 2 }
Console.WriteLine((int)Status.Active);
Console.WriteLine((int)Status.Inactive);
Console.WriteLine((int)Status.Pending);
"#), &["1", "0", "2"]);
}

#[test] fn enum_tostring_parse() {
    assert_eq!(run_csharp(r#"
enum Day { Mon, Tue, Wed, Thu, Fri }
Day d = Day.Wed;
Console.WriteLine(d.ToString());
"#), &["Wed"]);
}

#[test] fn enum_in_switch() {
    assert_eq!(run_csharp(r#"
enum Season { Spring, Summer, Autumn, Winter }
Season s = Season.Summer;
switch (s) {
    case Season.Spring: Console.WriteLine("spring"); break;
    case Season.Summer: Console.WriteLine("summer"); break;
    case Season.Autumn: Console.WriteLine("autumn"); break;
    case Season.Winter: Console.WriteLine("winter"); break;
}
"#), &["summer"]);
}

// ===================================================================
// CONST AND READONLY
// ===================================================================

#[test] fn const_field() {
    assert_eq!(run_csharp(r#"
class Config {
    public const int MaxRetries = 3;
    public const string AppName = "MyApp";
}
Console.WriteLine(Config.MaxRetries);
Console.WriteLine(Config.AppName);
"#), &["3", "MyApp"]);
}

#[test] fn readonly_field() {
    assert_eq!(run_csharp(r#"
class Circle {
    public readonly double Pi = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return Pi * Radius * Radius; }
}
var c = new Circle(1);
Console.WriteLine(c.Pi);
"#), &["3.14159"]);
}

// ===================================================================
// CASTING AND TYPE CONVERSION
// ===================================================================

#[test] fn implicit_cast() {
    assert_eq!(run_csharp(r#"
int i = 42;
double d = i;
Console.WriteLine(d);
"#), &["42"]);
}

#[test] fn explicit_cast() {
    assert_eq!(run_csharp(r#"
double d = 3.99;
int i = (int)d;
Console.WriteLine(i);
"#), &["3"]);
}

#[test] fn as_operator() {
    assert_eq!(run_csharp(r#"
object obj = "hello";
string s = obj as string;
Console.WriteLine(s != null ? s : "null");
int? i = obj as int?;
Console.WriteLine(i != null ? i.ToString() : "null");
"#), &["hello", "null"]);
}

// ===================================================================
// MATH OPERATIONS
// ===================================================================

#[test] fn math_abs_pow_sqrt() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(Math.Abs(-42));
Console.WriteLine(Math.Pow(2, 10));
Console.WriteLine(Math.Sqrt(144));
"#), &["42", "1024", "12"]);
}

#[test] fn math_min_max_round() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(Math.Min(3, 7));
Console.WriteLine(Math.Max(3, 7));
Console.WriteLine(Math.Round(3.7));
Console.WriteLine(Math.Floor(3.7));
Console.WriteLine(Math.Ceiling(3.2));
"#), &["3", "7", "4", "3", "4"]);
}

// ===================================================================
// ALGORITHMS
// ===================================================================

#[test] fn bubble_sort() {
    assert_eq!(run_csharp(r#"
int[] arr = { 5, 3, 8, 1, 2 };
for (int i = 0; i < arr.Length; i++) {
    for (int j = 0; j < arr.Length - 1 - i; j++) {
        if (arr[j] > arr[j + 1]) {
            int tmp = arr[j];
            arr[j] = arr[j + 1];
            arr[j + 1] = tmp;
        }
    }
}
Console.WriteLine(string.Join(",", arr));
"#), &["1,2,3,5,8"]);
}

#[test] fn binary_search_manual() {
    assert_eq!(run_csharp(r#"
int[] arr = { 1, 3, 5, 7, 9, 11, 13 };
int target = 7;
int lo = 0, hi = arr.Length - 1;
while (lo <= hi) {
    int mid = (lo + hi) / 2;
    if (arr[mid] == target) { Console.WriteLine("found at " + mid); break; }
    else if (arr[mid] < target) lo = mid + 1;
    else hi = mid - 1;
}
"#), &["found at 3"]);
}

#[test] fn reverse_string() {
    assert_eq!(run_csharp(r#"
string s = "Hello World";
char[] chars = s.ToCharArray();
Array.Reverse(chars);
Console.WriteLine(new string(chars));
"#), &["dlroW olleH"]);
}

#[test] fn fibonacci_iterative() {
    assert_eq!(run_csharp(r#"
int n = 10;
int a = 0, b = 1;
for (int i = 0; i < n; i++) {
    Console.WriteLine(a);
    int tmp = a + b;
    a = b;
    b = tmp;
}
"#), &["0", "1", "1", "2", "3", "5", "8", "13", "21", "34"]);
}

#[test] fn factorial_recursive() {
    assert_eq!(run_csharp(r#"
class Math2 {
    public static int Factorial(int n) {
        if (n <= 1) return 1;
        return n * Factorial(n - 1);
    }
}
Console.WriteLine(Math2.Factorial(0));
Console.WriteLine(Math2.Factorial(1));
Console.WriteLine(Math2.Factorial(5));
Console.WriteLine(Math2.Factorial(10));
"#), &["1", "1", "120", "3628800"]);
}

#[test] fn gcd_euclidean() {
    assert_eq!(run_csharp(r#"
class Algorithms {
    public static int GCD(int a, int b) {
        while (b != 0) { int t = b; b = a % b; a = t; }
        return a;
    }
}
Console.WriteLine(Algorithms.GCD(48, 18));
Console.WriteLine(Algorithms.GCD(100, 75));
"#), &["6", "25"]);
}

// ===================================================================
// FOR / FOREACH / WHILE / DO-WHILE
// ===================================================================

#[test] fn for_loop_with_break_continue() {
    assert_eq!(run_csharp(r#"
for (int i = 0; i < 10; i++) {
    if (i % 2 == 0) continue;
    if (i > 7) break;
    Console.WriteLine(i);
}
"#), &["1", "3", "5", "7"]);
}

#[test] fn while_loop() {
    assert_eq!(run_csharp(r#"
int x = 1;
while (x <= 16) {
    Console.WriteLine(x);
    x *= 2;
}
"#), &["1", "2", "4", "8", "16"]);
}

#[test] fn do_while_loop() {
    assert_eq!(run_csharp(r#"
int x = 1;
do {
    Console.WriteLine(x);
    x *= 3;
} while (x < 100);
"#), &["1", "3", "9", "27", "81"]);
}

// ===================================================================
// STATIC CLASSES AND METHODS
// ===================================================================

#[test] fn static_utility_class() {
    assert_eq!(run_csharp(r#"
static class StringUtils {
    public static bool IsPalindrome(string s) {
        string lower = s.ToLower();
        char[] chars = lower.ToCharArray();
        Array.Reverse(chars);
        return lower == new string(chars);
    }
}
Console.WriteLine(StringUtils.IsPalindrome("racecar"));
Console.WriteLine(StringUtils.IsPalindrome("hello"));
Console.WriteLine(StringUtils.IsPalindrome("Madam"));
"#), &["True", "False", "True"]);
}

// ===================================================================
// PROPERTIES WITH LOGIC
// ===================================================================

#[test] fn property_with_backing_field() {
    assert_eq!(run_csharp(r#"
class Temperature {
    private double celsius;
    public double Celsius {
        get { return celsius; }
        set {
            if (value < -273.15) celsius = -273.15;
            else celsius = value;
        }
    }
    public double Fahrenheit {
        get { return celsius * 9.0 / 5.0 + 32; }
    }
}
var t = new Temperature();
t.Celsius = 100;
Console.WriteLine(t.Fahrenheit);
t.Celsius = -500;
Console.WriteLine(t.Celsius);
"#), &["212", "-273.15"]);
}

// ===================================================================
// DICTIONARY PATTERNS
// ===================================================================

#[test] fn dictionary_word_count() {
    assert_eq!(run_csharp(r#"
string text = "the cat sat on the mat the cat";
var words = text.Split(' ');
var counts = new Dictionary<string, int>();
foreach (var w in words) {
    if (counts.ContainsKey(w)) counts[w]++;
    else counts[w] = 1;
}
Console.WriteLine("the: " + counts["the"]);
Console.WriteLine("cat: " + counts["cat"]);
Console.WriteLine("sat: " + counts["sat"]);
"#), &["the: 3", "cat: 2", "sat: 1"]);
}

#[test] fn dictionary_grouping() {
    assert_eq!(run_csharp(r#"
var data = new List<string> { "apple", "banana", "avocado", "blueberry", "cherry" };
var grouped = new Dictionary<char, List<string>>();
foreach (var item in data) {
    char key = item[0];
    if (!grouped.ContainsKey(key)) grouped[key] = new List<string>();
    grouped[key].Add(item);
}
Console.WriteLine("a: " + grouped['a'].Count);
Console.WriteLine("b: " + grouped['b'].Count);
Console.WriteLine("c: " + grouped['c'].Count);
"#), &["a: 2", "b: 2", "c: 1"]);
}
