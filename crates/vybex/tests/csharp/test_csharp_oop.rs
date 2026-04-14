use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: OOP — abstract, sealed, virtual/override, interfaces,
// structs, enums, records, constructor chaining
// ═══════════════════════════════════════════════════════════

#[test]
fn abstract_class_and_override() {
    let out = run_csharp(r#"
abstract class Shape {
    public abstract double Area();
    public string Describe() { return "Area=" + Area(); }
}
class Square : Shape {
    public double Side;
    public Square(double s) { Side = s; }
    public override double Area() { return Side * Side; }
}
var sq = new Square(5);
Console.WriteLine(sq.Area());
Console.WriteLine(sq.Describe());
"#);
    assert_eq!(out, vec!["25", "Area=25"]);
}

#[test]
fn sealed_class() {
    let out = run_csharp(r#"
sealed class Singleton {
    public int Value = 42;
}
var s = new Singleton();
Console.WriteLine(s.Value);
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn interface_with_multiple_methods() {
    let out = run_csharp(r#"
interface ICalculator {
    int Add(int a, int b);
    int Multiply(int a, int b);
}
class Calc : ICalculator {
    public int Add(int a, int b) { return a + b; }
    public int Multiply(int a, int b) { return a * b; }
}
var c = new Calc();
Console.WriteLine(c.Add(3, 4));
Console.WriteLine(c.Multiply(3, 4));
"#);
    assert_eq!(out, vec!["7", "12"]);
}

#[test]
fn struct_basic() {
    let out = run_csharp(r#"
struct Point {
    public int X;
    public int Y;
    public Point(int x, int y) { X = x; Y = y; }
    public int Sum() { return X + Y; }
}
var p = new Point(3, 4);
Console.WriteLine(p.Sum());
"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn enum_usage() {
    let out = run_csharp(r#"
enum Direction { North, South, East, West }
Direction d = Direction.East;
Console.WriteLine(d);
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_explicit_values() {
    let out = run_csharp(r#"
enum HttpStatus {
    OK = 200,
    NotFound = 404,
    ServerError = 500
}
Console.WriteLine(HttpStatus.NotFound);
"#);
    assert_eq!(out, vec!["404"]);
}

#[test]
fn record_with_properties() {
    let out = run_csharp(r#"
record Person(string Name, int Age);
var p = new Person("Alice", 30);
Console.WriteLine(p.Name);
Console.WriteLine(p.Age);
"#);
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn record_with_body() {
    let out = run_csharp(r#"
record Product(string Name, double Price) {
    public string Display() { return Name + ": $" + Price; }
}
var p = new Product("Widget", 9.99);
Console.WriteLine(p.Display());
"#);
    assert_eq!(out, vec!["Widget: $9.99"]);
}

#[test]
fn constructor_chaining_base() {
    let out = run_csharp(r#"
class Animal {
    public string Name;
    public Animal(string name) { Name = name; }
}
class Dog : Animal {
    public string Breed;
    public Dog(string name, string breed) : base(name) {
        Breed = breed;
    }
}
var d = new Dog("Rex", "Lab");
Console.WriteLine(d.Name);
Console.WriteLine(d.Breed);
"#);
    assert_eq!(out, vec!["Rex", "Lab"]);
}

#[test]
fn virtual_override_chain() {
    let out = run_csharp(r#"
class A {
    public virtual string Name() { return "A"; }
}
class B : A {
    public override string Name() { return "B"; }
}
var obj = new B();
Console.WriteLine(obj.Name());
"#);
    assert_eq!(out, vec!["B"]);
}

#[test]
fn class_with_property_getset() {
    let out = run_csharp(r#"
class Temperature {
    private double _celsius;
    public double Celsius {
        get { return _celsius; }
        set { _celsius = value; }
    }
    public double Fahrenheit {
        get { return _celsius * 9 / 5 + 32; }
    }
}
var t = new Temperature();
t.Celsius = 100;
Console.WriteLine(t.Celsius);
Console.WriteLine(t.Fahrenheit);
"#);
    assert_eq!(out, vec!["100", "212"]);
}

#[test]
fn class_auto_property_default() {
    let out = run_csharp(r#"
class Config {
    public string Name { get; set; } = "default";
    public int Count { get; set; } = 0;
}
var c = new Config();
Console.WriteLine(c.Name);
c.Name = "custom";
Console.WriteLine(c.Name);
"#);
    assert_eq!(out, vec!["default", "custom"]);
}

#[test]
fn class_with_static_field() {
    let out = run_csharp(r#"
class Counter {
    public static int Count = 0;
    public Counter() { Count++; }
}
var a = new Counter();
var b = new Counter();
var c = new Counter();
Console.WriteLine(Counter.Count);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn class_readonly_field() {
    let out = run_csharp(r#"
class Circle {
    public readonly double PI = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return PI * Radius * Radius; }
}
var c = new Circle(10);
Console.WriteLine(c.Area());
"#);
    assert_eq!(out, vec!["314.159"]);
}

#[test]
fn class_const_field() {
    let out = run_csharp(r#"
class MathConst {
    public const double PI = 3.14159;
    public const double E = 2.71828;
}
Console.WriteLine(MathConst.PI);
Console.WriteLine(MathConst.E);
"#);
    assert_eq!(out, vec!["3.14159", "2.71828"]);
}
