use super::helpers::run_csharp;

#[test]
fn class_with_constructor() {
    let out = run_csharp(
        r#"
        class Person {
            string name;
            int age;
            public Person(string n, int a) {
                this.name = n;
                this.age = a;
            }
            public string Describe() {
                return this.name + " is " + this.age;
            }
        }
        var p = new Person("Alice", 30);
        Console.WriteLine(p.Describe());
    "#,
    );
    assert_eq!(out, vec!["Alice is 30"]);
}

#[test]
fn class_field_access() {
    let out = run_csharp(
        r#"
        class Box {
            public int value;
            public Box(int v) { this.value = v; }
        }
        var b = new Box(42);
        Console.WriteLine(b.value);
    "#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn class_multiple_instances() {
    let out = run_csharp(
        r#"
        class Counter {
            int count;
            public Counter(int start) { this.count = start; }
            public void Inc() { this.count = this.count + 1; }
            public int Get() { return this.count; }
        }
        var a = new Counter(0);
        var b = new Counter(100);
        a.Inc(); a.Inc();
        b.Inc();
        Console.WriteLine(a.Get());
        Console.WriteLine(b.Get());
    "#,
    );
    assert_eq!(out, vec!["2", "101"]);
}

#[test]
fn inheritance_basic() {
    let out = run_csharp(
        r#"
        class Animal {
            string species;
            public Animal(string s) { this.species = s; }
            public string GetSpecies() { return this.species; }
        }
        class Dog : Animal {
            public Dog() : base("Canine") {}
        }
        var d = new Dog();
        Console.WriteLine(d.GetSpecies());
    "#,
    );
    assert_eq!(out, vec!["Canine"]);
}

#[test]
fn inheritance_override_method() {
    let out = run_csharp(
        r#"
        class Animal {
            string name;
            public Animal(string n) { this.name = n; }
            public string Speak() { return this.name + " speaks"; }
        }
        class Dog : Animal {
            public Dog(string n) : base(n) {}
            public string Bark() { return this.name + " barks"; }
        }
        var d = new Dog("Rex");
        Console.WriteLine(d.Speak());
        Console.WriteLine(d.Bark());
    "#,
    );
    assert_eq!(out, vec!["Rex speaks", "Rex barks"]);
}

#[test]
fn property_chain() {
    let out = run_csharp(
        r#"
        class Inner { public int value; public Inner(int v) { this.value = v; } }
        class Outer { public Inner inner; public Outer(int v) { this.inner = new Inner(v); } }
        var o = new Outer(42);
        Console.WriteLine(o.inner.value);
    "#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn two_classes_interacting() {
    let out = run_csharp(
        r#"
        class Engine {
            public int hp;
            public Engine(int h) { this.hp = h; }
        }
        class Car {
            public string model;
            public Engine engine;
            public Car(string m, int hp) {
                this.model = m;
                this.engine = new Engine(hp);
            }
        }
        var car = new Car("Tesla", 450);
        Console.WriteLine(car.model);
        Console.WriteLine(car.engine.hp);
    "#,
    );
    assert_eq!(out, vec!["Tesla", "450"]);
}

#[test]
fn linked_list_node() {
    let out = run_csharp(
        r#"
        class Node {
            public int value;
            public Node next;
            public Node(int v) { this.value = v; this.next = null; }
        }
        var a = new Node(1);
        var b = new Node(2);
        var c = new Node(3);
        a.next = b;
        b.next = c;
        Console.WriteLine(a.value);
        Console.WriteLine(a.next.value);
        Console.WriteLine(a.next.next.value);
    "#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn auto_property() {
    let out = run_csharp(
        r#"
        class Person {
            public string Name { get; set; }
            public Person(string n) { this.Name = n; }
        }
        var p = new Person("Alice");
        Console.WriteLine(p.Name);
    "#,
    );
    assert_eq!(out, vec!["Alice"]);
}

#[test]
fn auto_property_multiple() {
    let out = run_csharp(
        r#"
        class Car {
            public string Model { get; set; }
            public int Year { get; set; }
            public Car(string m, int y) { this.Model = m; this.Year = y; }
        }
        var c = new Car("Tesla", 2024);
        Console.WriteLine(c.Model);
        Console.WriteLine(c.Year);
    "#,
    );
    assert_eq!(out, vec!["Tesla", "2024"]);
}

#[test]
fn static_method_in_class() {
    let out = run_csharp(
        r#"
        class MathUtils {
            public static int Add(int a, int b) { return a + b; }
        }
        Console.WriteLine(MathUtils.Add(3, 4));
    "#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn recursive_factorial() {
    let out = run_csharp(
        r#"
        class MathUtils {
            public static int Fact(int n) {
                if (n <= 1) return 1;
                return Fact(n - 1) * n;
            }
        }
        Console.WriteLine(MathUtils.Fact(5));
    "#,
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn static_class_method_call() {
    let out = run_csharp(
        r#"
        class MathHelper {
            public static int Square(int x) { return x * x; }
            public static int Double(int x) { return x * 2; }
        }
        Console.WriteLine(MathHelper.Square(5));
        Console.WriteLine(MathHelper.Double(7));
    "#,
    );
    assert_eq!(out, vec!["25", "14"]);
}

#[test]
fn interface_basic() {
    let out = run_csharp(
        r#"
        interface IGreeter {
            string Greet();
        }
        class HelloGreeter : IGreeter {
            public string Greet() {
                return "Hello from interface!";
            }
        }
        var g = new HelloGreeter();
        Console.WriteLine(g.Greet());
    "#,
    );
    assert_eq!(out, vec!["Hello from interface!"]);
}

#[test]
fn enum_values() {
    let out = run_csharp(
        r#"
        enum Color { Red, Green, Blue }
        Console.WriteLine(Color.Red);
        Console.WriteLine(Color.Green);
        Console.WriteLine(Color.Blue);
    "#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn enum_explicit_values() {
    let out = run_csharp(
        r#"
        enum Status { Ok = 200, NotFound = 404, Error = 500 }
        Console.WriteLine(Status.Ok);
        Console.WriteLine(Status.NotFound);
    "#,
    );
    assert_eq!(out, vec!["200", "404"]);
}

#[test]
fn record_basic() {
    let out = run_csharp(
        r#"
        record Person(string Name, int Age);
        var p = new Person("Alice", 30);
        Console.WriteLine(p.Name);
        Console.WriteLine(p.Age);
    "#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn using_statement_basic() {
    let out = run_csharp(
        r#"
        class Resource {
            public string name;
            public Resource(string n) { this.name = n; }
        }
        using (var r = new Resource("test")) {
            Console.WriteLine(r.name);
        }
    "#,
    );
    assert_eq!(out, vec!["test"]);
}

#[test]
fn delegate_declaration_parses() {
    // Just verify delegate declaration parses correctly
    let out = run_csharp(
        r#"
        delegate int MathOp(int a, int b);
        Console.WriteLine("parsed");
    "#,
    );
    assert_eq!(out, vec!["parsed"]);
}
