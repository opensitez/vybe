use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Classes — constructors, methods, properties,
// inheritance, interfaces, static, abstract
// ═══════════════════════════════════════════════════════════

#[test]
fn class_basic() {
    let out = run_csharp(r#"
class Person {
    public string Name;
    public int Age;
    public Person(string name, int age) {
        Name = name;
        Age = age;
    }
    public string Describe() {
        return Name + " is " + Age;
    }
}
var p = new Person("Alice", 30);
Console.WriteLine(p.Describe());
"#);
    assert_eq!(out, vec!["Alice is 30"]);
}

#[test]
fn class_auto_property() {
    let out = run_csharp(r#"
class Config {
    public string Name { get; set; }
    public int Value { get; set; }
}
var c = new Config();
c.Name = "test";
c.Value = 42;
Console.WriteLine(c.Name);
Console.WriteLine(c.Value);
"#);
    assert_eq!(out, vec!["test", "42"]);
}

#[test]
fn class_inheritance() {
    let out = run_csharp(r#"
class Animal {
    public string Name;
    public Animal(string name) { Name = name; }
    public virtual string Speak() { return Name + " speaks"; }
}
class Dog : Animal {
    public Dog(string name) : base(name) {}
    public override string Speak() { return Name + " barks"; }
}
var d = new Dog("Rex");
Console.WriteLine(d.Speak());
"#);
    assert_eq!(out, vec!["Rex barks"]);
}

#[test]
fn class_super_call() {
    let out = run_csharp(r#"
class Base {
    public virtual string Greet() { return "Hello"; }
}
class Derived : Base {
    public override string Greet() { return base.Greet() + " World"; }
}
var d = new Derived();
Console.WriteLine(d.Greet());
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn class_static_method() {
    let out = run_csharp(r#"
class MathUtils {
    public static int Square(int x) { return x * x; }
}
Console.WriteLine(MathUtils.Square(7));
"#);
    assert_eq!(out, vec!["49"]);
}

#[test]
fn class_this_reference() {
    let out = run_csharp(r#"
class Counter {
    private int count = 0;
    public void Increment() { this.count++; }
    public int GetCount() { return this.count; }
}
var c = new Counter();
c.Increment();
c.Increment();
c.Increment();
Console.WriteLine(c.GetCount());
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn class_multiple_instances() {
    let out = run_csharp(r#"
class Box {
    public int Value;
    public Box(int v) { Value = v; }
}
var a = new Box(10);
var b = new Box(20);
Console.WriteLine(a.Value + b.Value);
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn interface_implementation() {
    let out = run_csharp(r#"
interface IShape {
    double Area();
}
class Circle : IShape {
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return 3.14159 * Radius * Radius; }
}
var c = new Circle(5);
Console.WriteLine(c.Area());
"#);
    assert_eq!(out, vec!["78.53975"]);
}

#[test]
fn enum_basic() {
    let out = run_csharp(r#"
enum Color { Red, Green, Blue }
Color c = Color.Green;
Console.WriteLine((int)c);
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn class_method_chaining() {
    let out = run_csharp(r#"
class Builder {
    private string result = "";
    public Builder Add(string s) {
        result += s;
        return this;
    }
    public string Build() { return result; }
}
var r = new Builder().Add("a").Add("b").Add("c").Build();
Console.WriteLine(r);
"#);
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn class_tostring() {
    let out = run_csharp(r#"
class Point {
    public int X;
    public int Y;
    public Point(int x, int y) { X = x; Y = y; }
    public override string ToString() { return "(" + X + ", " + Y + ")"; }
}
var p = new Point(3, 4);
Console.WriteLine(p.ToString());
"#);
    assert_eq!(out, vec!["(3, 4)"]);
}

#[test]
fn class_pass_by_reference() {
    let out = run_csharp(r#"
class Box {
    public int Value;
}
void Modify(Box b) {
    b.Value = 99;
}
var b = new Box();
b.Value = 1;
Modify(b);
Console.WriteLine(b.Value);
"#);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn recursive_factorial() {
    let out = run_csharp(r#"
int Factorial(int n) {
    if (n <= 1) return 1;
    return n * Factorial(n - 1);
}
Console.WriteLine(Factorial(6));
"#);
    assert_eq!(out, vec!["720"]);
}

#[test]
fn multi_level_inheritance() {
    let out = run_csharp(r#"
class A {
    public virtual string Who() { return "A"; }
}
class B : A {
    public override string Who() { return "B->" + base.Who(); }
}
class C : B {
    public override string Who() { return "C->" + base.Who(); }
}
var c = new C();
Console.WriteLine(c.Who());
"#);
    assert_eq!(out, vec!["C->B->A"]);
}
