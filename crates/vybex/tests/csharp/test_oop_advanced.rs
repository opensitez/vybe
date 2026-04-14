/// C# OOP patterns: abstract classes, sealed, virtual/override chains,
/// partial classes, static classes, nested classes, indexers, operator overloading,
/// explicit interface implementation, method hiding (new keyword).

use super::helpers::run_csharp;

// ===================================================================
// ABSTRACT CLASSES
// ===================================================================

#[test] fn abstract_class_basic() {
    assert_eq!(run_csharp(r#"
abstract class Shape {
    public abstract double Area();
    public string Describe() { return "I am a shape"; }
}
class Circle : Shape {
    double radius;
    public Circle(double r) { radius = r; }
    public override double Area() { return 3.14 * radius * radius; }
}
var c = new Circle(5);
Console.WriteLine(c.Area());
Console.WriteLine(c.Describe());
"#), &["78.5", "I am a shape"]);
}

#[test] fn abstract_multiple_implementations() {
    assert_eq!(run_csharp(r#"
abstract class Vehicle {
    public abstract string Type();
}
class Car : Vehicle {
    public override string Type() { return "Car"; }
}
class Truck : Vehicle {
    public override string Type() { return "Truck"; }
}
Vehicle v = new Car();
Console.WriteLine(v.Type());
v = new Truck();
Console.WriteLine(v.Type());
"#), &["Car", "Truck"]);
}

#[test] fn abstract_with_constructor() {
    assert_eq!(run_csharp(r#"
abstract class Animal {
    protected string name;
    public Animal(string n) { name = n; }
    public abstract string Sound();
    public string Greet() { return name + " says " + Sound(); }
}
class Dog : Animal {
    public Dog(string n) : base(n) { }
    public override string Sound() { return "Woof"; }
}
var d = new Dog("Rex");
Console.WriteLine(d.Greet());
"#), &["Rex says Woof"]);
}

// ===================================================================
// SEALED CLASSES
// ===================================================================

#[test] fn sealed_class_basic() {
    assert_eq!(run_csharp(r#"
sealed class Config {
    public string Name { get; set; }
    public Config(string n) { Name = n; }
}
var c = new Config("prod");
Console.WriteLine(c.Name);
"#), &["prod"]);
}

// ===================================================================
// VIRTUAL / OVERRIDE CHAINS
// ===================================================================

#[test] fn virtual_override_three_levels() {
    assert_eq!(run_csharp(r#"
class A {
    public virtual string Who() { return "A"; }
}
class B : A {
    public override string Who() { return "B"; }
}
class C : B {
    public override string Who() { return "C"; }
}
A obj = new C();
Console.WriteLine(obj.Who());
"#), &["C"]);
}

#[test] fn virtual_base_call() {
    assert_eq!(run_csharp(r#"
class Base {
    public virtual string Greet() { return "Hello"; }
}
class Child : Base {
    public override string Greet() { return base.Greet() + " World"; }
}
var c = new Child();
Console.WriteLine(c.Greet());
"#), &["Hello World"]);
}

// ===================================================================
// METHOD HIDING (new KEYWORD)
// ===================================================================

#[test] fn method_hiding_new() {
    assert_eq!(run_csharp(r#"
class Base {
    public string Speak() { return "base"; }
}
class Child : Base {
    public new string Speak() { return "child"; }
}
var c = new Child();
Console.WriteLine(c.Speak());
"#), &["child"]);
}

// ===================================================================
// STATIC CLASSES
// ===================================================================

#[test] fn static_class_methods() {
    assert_eq!(run_csharp(r#"
static class MathHelper {
    public static int Square(int x) { return x * x; }
    public static int Double(int x) { return x * 2; }
}
Console.WriteLine(MathHelper.Square(5));
Console.WriteLine(MathHelper.Double(7));
"#), &["25", "14"]);
}

#[test] fn static_class_with_constants() {
    assert_eq!(run_csharp(r#"
static class Constants {
    public const double Pi = 3.14159;
    public const int MaxSize = 100;
}
Console.WriteLine(Constants.Pi);
Console.WriteLine(Constants.MaxSize);
"#), &["3.14159", "100"]);
}

// ===================================================================
// NESTED CLASSES
// ===================================================================

#[test] fn nested_class_basic() {
    assert_eq!(run_csharp(r#"
class Outer {
    public class Inner {
        public string Hello() { return "inner"; }
    }
}
var i = new Outer.Inner();
Console.WriteLine(i.Hello());
"#), &["inner"]);
}

// ===================================================================
// INDEXERS
// ===================================================================

#[test] fn indexer_basic() {
    assert_eq!(run_csharp(r#"
class Sentence {
    string[] words;
    public Sentence(string[] w) { words = w; }
    public string this[int index] {
        get { return words[index]; }
        set { words[index] = value; }
    }
}
var s = new Sentence(new string[] { "hello", "world" });
Console.WriteLine(s[0]);
Console.WriteLine(s[1]);
s[1] = "C#";
Console.WriteLine(s[1]);
"#), &["hello", "world", "C#"]);
}

// ===================================================================
// OPERATOR OVERLOADING
// ===================================================================

#[test] fn operator_overload_plus() {
    assert_eq!(run_csharp(r#"
class Vector {
    public double X { get; set; }
    public double Y { get; set; }
    public Vector(double x, double y) { X = x; Y = y; }
    public static Vector operator +(Vector a, Vector b) {
        return new Vector(a.X + b.X, a.Y + b.Y);
    }
}
var a = new Vector(1, 2);
var b = new Vector(3, 4);
var c = a + b;
Console.WriteLine(c.X);
Console.WriteLine(c.Y);
"#), &["4", "6"]);
}

#[test] fn operator_overload_equals() {
    assert_eq!(run_csharp(r#"
class Point {
    public int X { get; set; }
    public int Y { get; set; }
    public Point(int x, int y) { X = x; Y = y; }
    public static bool operator ==(Point a, Point b) {
        return a.X == b.X && a.Y == b.Y;
    }
    public static bool operator !=(Point a, Point b) {
        return !(a == b);
    }
}
var a = new Point(1, 2);
var b = new Point(1, 2);
var c = new Point(3, 4);
Console.WriteLine(a == b);
Console.WriteLine(a != c);
"#), &["True", "True"]);
}

// ===================================================================
// OBJECT COMPOSITION
// ===================================================================

#[test] fn composition_engine_car() {
    assert_eq!(run_csharp(r#"
class Engine {
    public int Horsepower { get; set; }
    public Engine(int hp) { Horsepower = hp; }
}
class Car {
    public string Name { get; set; }
    public Engine Engine { get; set; }
    public Car(string name, int hp) {
        Name = name;
        Engine = new Engine(hp);
    }
    public string Info() { return Name + " " + Engine.Horsepower + "hp"; }
}
var car = new Car("Sedan", 200);
Console.WriteLine(car.Info());
"#), &["Sedan 200hp"]);
}

// ===================================================================
// CONSTRUCTOR CHAINING
// ===================================================================

#[test] fn constructor_chaining_this() {
    assert_eq!(run_csharp(r#"
class Point {
    public int X { get; set; }
    public int Y { get; set; }
    public Point() : this(0, 0) { }
    public Point(int x, int y) { X = x; Y = y; }
}
var a = new Point();
var b = new Point(5, 10);
Console.WriteLine(a.X + "," + a.Y);
Console.WriteLine(b.X + "," + b.Y);
"#), &["0,0", "5,10"]);
}

// ===================================================================
// READONLY PROPERTIES
// ===================================================================

#[test] fn readonly_auto_property() {
    assert_eq!(run_csharp(r#"
class Person {
    public string Name { get; }
    public Person(string name) { Name = name; }
}
var p = new Person("Alice");
Console.WriteLine(p.Name);
"#), &["Alice"]);
}

// ===================================================================
// EXPRESSION-BODIED MEMBERS
// ===================================================================

#[test] fn expression_bodied_method() {
    assert_eq!(run_csharp(r#"
class Calc {
    public int Square(int x) => x * x;
    public string Greet(string name) => "Hello " + name;
}
var c = new Calc();
Console.WriteLine(c.Square(7));
Console.WriteLine(c.Greet("World"));
"#), &["49", "Hello World"]);
}

#[test] fn expression_bodied_property() {
    assert_eq!(run_csharp(r#"
class Circle {
    public double Radius { get; set; }
    public double Area => 3.14 * Radius * Radius;
    public Circle(double r) { Radius = r; }
}
var c = new Circle(5);
Console.WriteLine(c.Area);
"#), &["78.5"]);
}

// ===================================================================
// POLYMORPHIC COLLECTIONS
// ===================================================================

#[test] fn polymorphic_list() {
    assert_eq!(run_csharp(r#"
class Animal {
    public virtual string Speak() { return "..."; }
}
class Dog : Animal {
    public override string Speak() { return "Woof"; }
}
class Cat : Animal {
    public override string Speak() { return "Meow"; }
}
var animals = new List<Animal> { new Dog(), new Cat(), new Dog() };
foreach (var a in animals) {
    Console.WriteLine(a.Speak());
}
"#), &["Woof", "Meow", "Woof"]);
}

// ===================================================================
// TOSTRING OVERRIDE
// ===================================================================

#[test] fn tostring_override() {
    assert_eq!(run_csharp(r#"
class Person {
    public string Name { get; set; }
    public int Age { get; set; }
    public Person(string name, int age) { Name = name; Age = age; }
    public override string ToString() { return Name + " (" + Age + ")"; }
}
var p = new Person("Alice", 30);
Console.WriteLine(p.ToString());
Console.WriteLine(p);
"#), &["Alice (30)", "Alice (30)"]);
}

// ===================================================================
// THIS REFERENCE
// ===================================================================

#[test] fn this_reference_return() {
    assert_eq!(run_csharp(r#"
class Builder {
    string parts = "";
    public Builder Add(string part) {
        if (parts.Length > 0) parts += ", ";
        parts += part;
        return this;
    }
    public string Build() { return "[" + parts + "]"; }
}
var b = new Builder();
Console.WriteLine(b.Add("A").Add("B").Add("C").Build());
"#), &["[A, B, C]"]);
}
