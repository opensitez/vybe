/// C# interfaces, generics, generic constraints, multiple interface
/// implementation, IComparable, IEnumerable, covariance/contravariance,
/// default interface methods.
use super::helpers::run_csharp;

// ===================================================================
// INTERFACES
// ===================================================================

#[test]
fn interface_basic() {
    assert_eq!(
        run_csharp(
            r#"
interface IGreeter {
    string Greet();
}
class HelloGreeter : IGreeter {
    public string Greet() { return "Hello!"; }
}
IGreeter g = new HelloGreeter();
Console.WriteLine(g.Greet());
"#
        ),
        &["Hello!"]
    );
}

#[test]
fn interface_multiple_impl() {
    assert_eq!(
        run_csharp(
            r#"
interface IShape {
    double Area();
}
class Circle : IShape {
    public double Radius;
    public double Area() { return 3.14 * Radius * Radius; }
}
class Square : IShape {
    public double Side;
    public double Area() { return Side * Side; }
}
IShape c = new Circle { Radius = 10 };
IShape s = new Square { Side = 5 };
Console.WriteLine(c.Area());
Console.WriteLine(s.Area());
"#
        ),
        &["314", "25"]
    );
}

#[test]
fn multiple_interfaces() {
    assert_eq!(
        run_csharp(
            r#"
interface IPrintable {
    void Print();
}
interface ISerializable {
    string Serialize();
}
class Doc : IPrintable, ISerializable {
    public string Name;
    public void Print() { Console.WriteLine("Printing: " + Name); }
    public string Serialize() { return "DOC:" + Name; }
}
var d = new Doc { Name = "test" };
d.Print();
Console.WriteLine(d.Serialize());
"#
        ),
        &["Printing: test", "DOC:test"]
    );
}

#[test]
fn interface_property() {
    assert_eq!(
        run_csharp(
            r#"
interface INamed {
    string Name { get; }
}
class Person : INamed {
    public string Name { get; set; }
}
INamed p = new Person { Name = "Alice" };
Console.WriteLine(p.Name);
"#
        ),
        &["Alice"]
    );
}

#[test]
fn interface_polymorphic_list() {
    assert_eq!(
        run_csharp(
            r#"
interface IAnimal {
    string Speak();
}
class Dog : IAnimal {
    public string Speak() { return "Woof"; }
}
class Cat : IAnimal {
    public string Speak() { return "Meow"; }
}
var animals = new List<IAnimal> { new Dog(), new Cat(), new Dog() };
foreach (var a in animals) Console.WriteLine(a.Speak());
"#
        ),
        &["Woof", "Meow", "Woof"]
    );
}

#[test]
fn interface_is_check() {
    assert_eq!(
        run_csharp(
            r#"
interface IFlyable { }
class Bird : IFlyable { }
class Fish { }
object b = new Bird();
object f = new Fish();
Console.WriteLine(b is IFlyable);
Console.WriteLine(f is IFlyable);
"#
        ),
        &["True", "False"]
    );
}

// ===================================================================
// GENERICS
// ===================================================================

#[test]
fn generic_class() {
    assert_eq!(
        run_csharp(
            r#"
class Box<T> {
    public T Value;
    public Box(T val) { Value = val; }
}
var intBox = new Box<int>(42);
var strBox = new Box<string>("hello");
Console.WriteLine(intBox.Value);
Console.WriteLine(strBox.Value);
"#
        ),
        &["42", "hello"]
    );
}

#[test]
fn generic_method() {
    assert_eq!(
        run_csharp(
            r#"
class Utils {
    public static T Max<T>(T a, T b) where T : IComparable<T> {
        return a.CompareTo(b) > 0 ? a : b;
    }
}
Console.WriteLine(Utils.Max(3, 7));
Console.WriteLine(Utils.Max("apple", "banana"));
"#
        ),
        &["7", "banana"]
    );
}

#[test]
fn generic_multiple_type_params() {
    assert_eq!(
        run_csharp(
            r#"
class Pair<TFirst, TSecond> {
    public TFirst First;
    public TSecond Second;
    public Pair(TFirst f, TSecond s) { First = f; Second = s; }
    public override string ToString() { return First + ":" + Second; }
}
var p = new Pair<string, int>("age", 30);
Console.WriteLine(p);
"#
        ),
        &["age:30"]
    );
}

#[test]
fn generic_interface() {
    assert_eq!(
        run_csharp(
            r#"
interface IRepository<T> {
    void Add(T item);
    int Count();
}
class ListRepo<T> : IRepository<T> {
    private List<T> items = new List<T>();
    public void Add(T item) { items.Add(item); }
    public int Count() { return items.Count; }
}
var repo = new ListRepo<string>();
repo.Add("a");
repo.Add("b");
repo.Add("c");
Console.WriteLine(repo.Count());
"#
        ),
        &["3"]
    );
}

#[test]
fn generic_stack_implementation() {
    assert_eq!(
        run_csharp(
            r#"
class MyStack<T> {
    private List<T> items = new List<T>();
    public void Push(T item) { items.Add(item); }
    public T Pop() {
        T item = items[items.Count - 1];
        items.RemoveAt(items.Count - 1);
        return item;
    }
    public int Count { get { return items.Count; } }
}
var s = new MyStack<int>();
s.Push(10);
s.Push(20);
s.Push(30);
Console.WriteLine(s.Pop());
Console.WriteLine(s.Pop());
Console.WriteLine(s.Count);
"#
        ),
        &["30", "20", "1"]
    );
}

// ===================================================================
// GENERIC CONSTRAINTS
// ===================================================================

#[test]
fn generic_where_new() {
    assert_eq!(
        run_csharp(
            r#"
class Factory<T> where T : new() {
    public T Create() { return new T(); }
}
class Item {
    public string Name = "default";
}
var f = new Factory<Item>();
var item = f.Create();
Console.WriteLine(item.Name);
"#
        ),
        &["default"]
    );
}

#[test]
fn generic_where_class_constraint() {
    assert_eq!(
        run_csharp(
            r#"
class Container<T> where T : class {
    public T Value;
    public bool IsNull() { return Value == null; }
}
var c = new Container<string>();
Console.WriteLine(c.IsNull());
c.Value = "hello";
Console.WriteLine(c.IsNull());
"#
        ),
        &["True", "False"]
    );
}

// ===================================================================
// ICOMPARABLE
// ===================================================================

#[test]
fn icomparable_implementation() {
    assert_eq!(
        run_csharp(
            r#"
class Temperature : IComparable<Temperature> {
    public double Degrees;
    public Temperature(double d) { Degrees = d; }
    public int CompareTo(Temperature other) {
        return Degrees.CompareTo(other.Degrees);
    }
    public override string ToString() { return Degrees + "°"; }
}
var temps = new List<Temperature> {
    new Temperature(100),
    new Temperature(37),
    new Temperature(0)
};
temps.Sort();
foreach (var t in temps) Console.WriteLine(t);
"#
        ),
        &["0°", "37°", "100°"]
    );
}

// ===================================================================
// IENUMERABLE / YIELD RETURN
// ===================================================================

#[test]
fn yield_return_basic() {
    assert_eq!(
        run_csharp(
            r#"
class Numbers {
    public static IEnumerable<int> OneToFive() {
        yield return 1;
        yield return 2;
        yield return 3;
        yield return 4;
        yield return 5;
    }
}
foreach (var n in Numbers.OneToFive()) Console.WriteLine(n);
"#
        ),
        &["1", "2", "3", "4", "5"]
    );
}

#[test]
fn yield_return_with_logic() {
    assert_eq!(
        run_csharp(
            r#"
class Gen {
    public static IEnumerable<int> EvenNumbers(int max) {
        for (int i = 0; i <= max; i++) {
            if (i % 2 == 0) yield return i;
        }
    }
}
foreach (var n in Gen.EvenNumbers(10)) Console.WriteLine(n);
"#
        ),
        &["0", "2", "4", "6", "8", "10"]
    );
}

#[test]
fn yield_return_fibonacci() {
    assert_eq!(
        run_csharp(
            r#"
class Fib {
    public static IEnumerable<int> Sequence(int count) {
        int a = 0, b = 1;
        for (int i = 0; i < count; i++) {
            yield return a;
            int temp = a + b;
            a = b;
            b = temp;
        }
    }
}
foreach (var n in Fib.Sequence(8)) Console.WriteLine(n);
"#
        ),
        &["0", "1", "1", "2", "3", "5", "8", "13"]
    );
}

// ===================================================================
// EXTENSION METHODS
// ===================================================================

#[test]
fn extension_method_basic() {
    assert_eq!(
        run_csharp(
            r#"
static class StringExtensions {
    public static string Reverse(this string s) {
        char[] chars = s.ToCharArray();
        Array.Reverse(chars);
        return new string(chars);
    }
}
Console.WriteLine("hello".Reverse());
"#
        ),
        &["olleh"]
    );
}

#[test]
fn extension_method_on_int() {
    assert_eq!(
        run_csharp(
            r#"
static class IntExtensions {
    public static bool IsEven(this int n) { return n % 2 == 0; }
    public static int Square(this int n) { return n * n; }
}
Console.WriteLine(4.IsEven());
Console.WriteLine(3.IsEven());
Console.WriteLine(5.Square());
"#
        ),
        &["True", "False", "25"]
    );
}

#[test]
fn extension_method_on_list() {
    assert_eq!(
        run_csharp(
            r#"
static class ListExtensions {
    public static string Join<T>(this List<T> list, string sep) {
        return string.Join(sep, list);
    }
}
var nums = new List<int> { 1, 2, 3, 4, 5 };
Console.WriteLine(nums.Join(", "));
"#
        ),
        &["1, 2, 3, 4, 5"]
    );
}
