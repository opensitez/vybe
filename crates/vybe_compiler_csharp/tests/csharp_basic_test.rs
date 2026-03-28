use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value};

fn run_cs(source: &str) -> Vec<String> {
    let unit = vybe_parser_csharp::parse(source).unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybe_host::setup_namespaces(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    let chunks = vybe_compiler_csharp::Compiler::new().compile(&unit)
        .unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));
    let result = output.borrow().clone();
    result
}

fn run_cs_one(source: &str) -> String {
    run_cs(source).into_iter().next().unwrap_or_default()
}

// ============================================================
// HELLO WORLD
// ============================================================

#[test]
fn hello_world() {
    let out = run_cs(r#"
        Console.WriteLine("Hello, World!");
    "#);
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn console_writeline_number() {
    assert_eq!(run_cs_one("Console.WriteLine(42);"), "42");
}

#[test]
fn console_writeline_bool() {
    assert_eq!(run_cs_one("Console.WriteLine(true);"), "true");
}

// ============================================================
// VARIABLES AND EXPRESSIONS
// ============================================================

#[test]
fn var_declaration() {
    let out = run_cs(r#"
        var x = 10;
        var y = 20;
        Console.WriteLine(x + y);
    "#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn typed_declaration() {
    let out = run_cs(r#"
        int x = 5;
        double y = 3.14;
        Console.WriteLine(x);
        Console.WriteLine(y);
    "#);
    assert_eq!(out, vec!["5", "3.14"]);
}

#[test]
fn string_concat() {
    assert_eq!(run_cs_one(r#"Console.WriteLine("hello" + " " + "world");"#), "hello world");
}

#[test]
fn arithmetic() {
    assert_eq!(run_cs_one("Console.WriteLine(2 + 3 * 4);"), "14");
}

#[test]
fn comparison() {
    assert_eq!(run_cs_one("Console.WriteLine(5 > 3);"), "true");
    assert_eq!(run_cs_one("Console.WriteLine(5 < 3);"), "false");
}

// ============================================================
// CONTROL FLOW
// ============================================================

#[test]
fn if_else() {
    let out = run_cs(r#"
        var x = 10;
        if (x > 5) {
            Console.WriteLine("big");
        } else {
            Console.WriteLine("small");
        }
    "#);
    assert_eq!(out, vec!["big"]);
}

#[test]
fn for_loop() {
    let out = run_cs(r#"
        var sum = 0;
        for (var i = 1; i <= 5; i++) {
            sum = sum + i;
        }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn while_loop() {
    let out = run_cs(r#"
        var i = 0;
        while (i < 3) {
            i = i + 1;
        }
        Console.WriteLine(i);
    "#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn foreach_array() {
    let out = run_cs(r#"
        var sum = 0;
        foreach (var x in new int[] { 10, 20, 30 }) {
            sum = sum + x;
        }
        Console.WriteLine(sum);
    "#);
    assert_eq!(out, vec!["60"]);
}

// ============================================================
// FUNCTIONS
// ============================================================

#[test]
fn static_method_in_class() {
    let out = run_cs(r#"
        class Program {
            static int Add(int a, int b) { return a + b; }
            static void Main() {
                Console.WriteLine(Add(3, 4));
            }
        }
    "#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn recursive_factorial() {
    let out = run_cs(r#"
        class Program {
            static int Fact(int n) {
                if (n <= 1) return 1;
                return n * Fact(n - 1);
            }
            static void Main() {
                Console.WriteLine(Fact(5));
            }
        }
    "#);
    assert_eq!(out, vec!["120"]);
}

// ============================================================
// CLASSES
// ============================================================

#[test]
fn class_with_constructor() {
    let out = run_cs(r#"
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
    "#);
    assert_eq!(out, vec!["Alice is 30"]);
}

#[test]
fn class_field_access() {
    let out = run_cs(r#"
        class Box {
            public int value;
            public Box(int v) { this.value = v; }
        }
        var b = new Box(42);
        Console.WriteLine(b.value);
    "#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn class_multiple_instances() {
    let out = run_cs(r#"
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
    "#);
    assert_eq!(out, vec!["2", "101"]);
}

// ============================================================
// INHERITANCE
// ============================================================

#[test]
fn inheritance_basic() {
    let out = run_cs(r#"
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
    "#);
    assert_eq!(out, vec!["Canine"]);
}

// ============================================================
// MATH
// ============================================================

#[test]
fn math_floor() {
    assert_eq!(run_cs_one("Console.WriteLine(Math.Floor(3.7));"), "3");
}

#[test]
fn math_abs() {
    assert_eq!(run_cs_one("Console.WriteLine(Math.Abs(-5));"), "5");
}

#[test]
fn math_sqrt() {
    assert_eq!(run_cs_one("Console.WriteLine(Math.Sqrt(16));"), "4");
}

// ============================================================
// STRINGS
// ============================================================

#[test]
fn string_length() {
    assert_eq!(run_cs_one(r#"Console.WriteLine("hello".Length);"#), "5");
}

#[test]
fn string_toupper() {
    assert_eq!(run_cs_one(r#"Console.WriteLine("hello".ToUpper());"#), "HELLO");
}

#[test]
fn string_contains() {
    assert_eq!(run_cs_one(r#"Console.WriteLine("hello world".Contains("world"));"#), "true");
}

// ============================================================
// ARRAYS
// ============================================================

#[test]
fn array_creation_and_index() {
    let out = run_cs(r#"
        var arr = new int[] { 10, 20, 30 };
        Console.WriteLine(arr[1]);
    "#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn array_length() {
    assert_eq!(run_cs_one("Console.WriteLine(new int[] { 1, 2, 3 }.Length);"), "3");
}

// ============================================================
// TRY/CATCH
// ============================================================

#[test]
fn try_catch_basic() {
    let out = run_cs(r#"
        try {
            throw new Exception("oops");
        } catch (Exception e) {
            Console.WriteLine("caught");
        }
    "#);
    assert_eq!(out, vec!["caught"]);
}

// ============================================================
// TERNARY
// ============================================================

#[test]
fn ternary_expression() {
    assert_eq!(run_cs_one("Console.WriteLine(5 > 3 ? \"yes\" : \"no\");"), "yes");
}

// ============================================================
// NULL COALESCING
// ============================================================

#[test]
fn null_coalescing() {
    assert_eq!(run_cs_one(r#"string s = null; Console.WriteLine(s ?? "default");"#), "default");
}
