use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Common patterns — builder, factory, singleton,
// composition, iteration, algorithms
// ═══════════════════════════════════════════════════════════

#[test]
fn builder_pattern() {
    let out = run_csharp(
        r#"
class QueryBuilder {
    private string query = "SELECT *";
    public QueryBuilder From(string table) {
        query += " FROM " + table;
        return this;
    }
    public QueryBuilder Where(string condition) {
        query += " WHERE " + condition;
        return this;
    }
    public string Build() { return query; }
}
var q = new QueryBuilder().From("users").Where("age > 18").Build();
Console.WriteLine(q);
"#,
    );
    assert_eq!(out, vec!["SELECT * FROM users WHERE age > 18"]);
}

#[test]
fn factory_method() {
    let out = run_csharp(
        r#"
class Shape {
    public string Type;
    private Shape(string t) { Type = t; }
    public static Shape Circle() { return new Shape("circle"); }
    public static Shape Square() { return new Shape("square"); }
}
var c = Shape.Circle();
var s = Shape.Square();
Console.WriteLine(c.Type);
Console.WriteLine(s.Type);
"#,
    );
    assert_eq!(out, vec!["circle", "square"]);
}

#[test]
fn fibonacci() {
    let out = run_csharp(
        r#"
class Program {
    public static int Fib(int n) {
        if (n <= 1) return n;
        return Fib(n - 1) + Fib(n - 2);
    }
}
Console.WriteLine(Program.Fib(10));
"#,
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn bubble_sort() {
    let out = run_csharp(
        r#"
var arr = new[] { 5, 3, 8, 1, 2 };
for (int i = 0; i < arr.Length - 1; i++) {
    for (int j = 0; j < arr.Length - 1 - i; j++) {
        if (arr[j] > arr[j + 1]) {
            int temp = arr[j];
            arr[j] = arr[j + 1];
            arr[j + 1] = temp;
        }
    }
}
foreach (var x in arr) Console.WriteLine(x);
"#,
    );
    assert_eq!(out, vec!["1", "2", "3", "5", "8"]);
}

#[test]
fn accumulator_pattern() {
    let out = run_csharp(
        r#"
var items = new[] { 1, 2, 3, 4, 5 };
int sum = 0;
int product = 1;
foreach (var x in items) {
    sum += x;
    product *= x;
}
Console.WriteLine(sum);
Console.WriteLine(product);
"#,
    );
    assert_eq!(out, vec!["15", "120"]);
}

#[test]
fn string_reversal() {
    let out = run_csharp(
        r#"
class StringUtils {
    public static string Reverse(string s) {
        string result = "";
        for (int i = s.Length - 1; i >= 0; i--) {
            result += s[i];
        }
        return result;
    }
}
Console.WriteLine(StringUtils.Reverse("hello"));
Console.WriteLine(StringUtils.Reverse("abcde"));
"#,
    );
    assert_eq!(out, vec!["olleh", "edcba"]);
}

#[test]
fn composition_with_classes() {
    let out = run_csharp(
        r#"
class Engine {
    public int Horsepower;
    public Engine(int hp) { Horsepower = hp; }
}
class Car {
    public string Make;
    public Engine Engine;
    public Car(string make, int hp) {
        Make = make;
        Engine = new Engine(hp);
    }
    public string Describe() {
        return Make + " " + Engine.Horsepower + "hp";
    }
}
var car = new Car("Toyota", 200);
Console.WriteLine(car.Describe());
"#,
    );
    assert_eq!(out, vec!["Toyota 200hp"]);
}

#[test]
fn multiple_classes_interacting() {
    let out = run_csharp(
        r#"
class Item {
    public string Name;
    public double Price;
    public Item(string n, double p) { Name = n; Price = p; }
}
class Cart {
    private List<Item> items = new List<Item>();
    public void Add(Item item) { items.Add(item); }
    public int Count() { return items.Count; }
    public double Total() {
        double sum = 0;
        foreach (var item in items) sum += item.Price;
        return sum;
    }
}
var cart = new Cart();
cart.Add(new Item("Apple", 1.5));
cart.Add(new Item("Bread", 2.5));
cart.Add(new Item("Milk", 3.0));
Console.WriteLine(cart.Count());
Console.WriteLine(cart.Total());
"#,
    );
    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn nested_class_access() {
    let out = run_csharp(
        r#"
class Outer {
    public int Value = 10;
    public class Inner {
        public int Value = 20;
    }
}
var o = new Outer();
var i = new Outer.Inner();
Console.WriteLine(o.Value);
Console.WriteLine(i.Value);
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn method_overloading() {
    let out = run_csharp(
        r#"
class Printer {
    public string Print(int x) { return "int:" + x; }
    public string Print(string x) { return "str:" + x; }
    public string Print(int x, int y) { return "pair:" + x + "," + y; }
}
var p = new Printer();
Console.WriteLine(p.Print(42));
Console.WriteLine(p.Print("hi"));
Console.WriteLine(p.Print(1, 2));
"#,
    );
    assert_eq!(out, vec!["int:42", "str:hi", "pair:1,2"]);
}
