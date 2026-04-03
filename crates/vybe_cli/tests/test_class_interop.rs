/// Comprehensive cross-language class interop tests.
/// These tests verify that classes defined in one language can be
/// instantiated, inherited, and used from any other language.
///
/// Currently: some tests are ignored pending JS/C# migration to common::classes.
/// After migration, all tests should pass.

use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value, HostContext};

fn setup_vm() -> (VM, Rc<RefCell<Vec<String>>>) {
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    (vm, output)
}

// ═══════════════════════════════════════════════════════════
// SECTION 1: CROSS-LANGUAGE CLASS INSTANTIATION
// ═══════════════════════════════════════════════════════════

#[test]
fn vb_class_js_instantiate() {
    let (mut vm, output) = setup_vm();
    let vb_src = r#"
Class Animal
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Function Speak() As String
        Return Name & " speaks"
    End Function
End Class
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile")).expect("VB run");

    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = "var a = new animal(\"Rex\"); console.log(a.speak());";
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

#[test]
fn vb_class_js_multiple_instances() {
    let (mut vm, output) = setup_vm();
    let vb_src = r#"
Class Counter
    Public Value As Integer
    Public Sub New(start As Integer)
        Value = start
    End Sub
    Public Sub Increment()
        Value = Value + 1
    End Sub
End Class
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile")).expect("VB run");

    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = r#"
var c1 = new counter(0);
var c2 = new counter(100);
c1.increment(); c1.increment(); c2.increment();
console.log(c1.value); console.log(c2.value);
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    assert_eq!(output.borrow().as_slice(), &["2", "101"]);
}

#[test]
fn js_class_cs_instantiate() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = r#"
class calculator {
    constructor() { this.result = 0; }
    add(n) { this.result = this.result + n; return this; }
    getResult() { return this.result; }
}
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    let cs_src = "var calc = new calculator(); calc.add(10); calc.add(20); Console.WriteLine(calc.getResult());";
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile")).expect("C# run");

    assert_eq!(output.borrow().as_slice(), &["30"]);
}

#[test]
fn cs_class_js_instantiate() {
    let (mut vm, output) = setup_vm();
    let cs_src = r#"
class Point {
    public int X; public int Y;
    public Point(int x, int y) { X = x; Y = y; }
}
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile")).expect("C# run");

    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = "var p = new point(10, 20); console.log(p.x + p.y);";
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    assert_eq!(output.borrow().as_slice(), &["30"]);
}

#[test]
fn vb_class_cs_instantiate() {
    let (mut vm, output) = setup_vm();
    let vb_src = r#"
Class Formatter
    Public Sub New()
    End Sub
    Public Function Format(text As String) As String
        Return "[" & text & "]"
    End Function
End Class
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile")).expect("VB run");

    let cs_src = "var f = new formatter(); Console.WriteLine(f.format(\"hello\"));";
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile")).expect("C# run");

    assert_eq!(output.borrow().as_slice(), &["[hello]"]);
}

#[test]
fn ruby_class_js_instantiate() {
    let (mut vm, output) = setup_vm();
    let rb_src = "class Dog\n  def initialize(name)\n    @name = name\n  end\n  def bark\n    @name\n  end\nend";
    let rb_prog = vybe_parser_ruby::parse(rb_src).expect("Ruby parse");
    vm.run(vybe_compiler_ruby::Compiler::new().compile(&rb_prog).expect("Ruby compile")).expect("Ruby run");

    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = "var d = new dog(\"Rex\"); console.log(d.bark());";
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    assert_eq!(output.borrow().as_slice(), &["Rex"]);
}

// ═══════════════════════════════════════════════════════════
// SECTION 2: CROSS-LANGUAGE INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn vb_parent_js_child_inherits() {
    let (mut vm, output) = setup_vm();
    let vb_src = r#"
Class Shape
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Function Area() As Integer
        Return 0
    End Function
End Class
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile")).expect("VB run");

    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = r#"
class rect extends shape {
    constructor(w, h) { super("rect"); this.w = w; this.h = h; }
    area() { return this.w * this.h; }
}
var r = new rect(5, 3);
console.log(r.name);
console.log(r.area());
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    assert_eq!(output.borrow().as_slice(), &["rect", "15"]);
}

#[test]
fn js_parent_cs_child_inherits() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = r#"
class vehicle {
    constructor(type) { this.type = type; this.speed = 0; }
    accelerate(n) { this.speed = this.speed + n; }
}
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    let cs_src = r#"
class Car : vehicle {
    public Car() : base("car") {}
}
var c = new Car();
c.accelerate(60);
Console.WriteLine(c.speed);
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile")).expect("C# run");

    assert_eq!(output.borrow().as_slice(), &["60"]);
}

// ═══════════════════════════════════════════════════════════
// SECTION 3: CROSS-LANGUAGE FUNCTION CALLS
// ═══════════════════════════════════════════════════════════

#[test]
fn php_function_js_calls() {
    let (mut vm, output) = setup_vm();
    let php_src = "<?php function double($x) { return $x * 2; } echo double(21);";
    let php_prog = vybe_parser_php::parse(php_src).expect("PHP parse");
    vm.run(vybe_compiler_php::Compiler::new().compile(&php_prog).expect("PHP compile")).expect("PHP run");

    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = "var r = double(50); console.log(r);";
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    let out = output.borrow();
    assert_eq!(out[0], "42");
    assert_eq!(out[1], "100");
}

#[test]
fn php_function_python_calls() {
    let (mut vm, output) = setup_vm();
    let php_src = "<?php function triple($x) { return $x * 3; }";
    let php_prog = vybe_parser_php::parse(php_src).expect("PHP parse");
    vm.run(vybe_compiler_php::Compiler::new().compile(&php_prog).expect("PHP compile")).expect("PHP run");

    let py_src = "result = triple(10)\nprint(result)";
    let py_prog = vybe_parser_python::parse(py_src).expect("Python parse");
    vm.run(vybe_compiler_python::Compiler::new().compile(&py_prog).expect("Python compile")).expect("Python run");

    assert_eq!(output.borrow().as_slice(), &["30"]);
}

#[test]
fn vb_function_ruby_calls() {
    let (mut vm, output) = setup_vm();
    let vb_src = r#"
Public Function Add(a As Integer, b As Integer) As Integer
    Return a + b
End Function
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile")).expect("VB run");

    let rb_src = "result = add(7, 8)\nputs result";
    let rb_prog = vybe_parser_ruby::parse(rb_src).expect("Ruby parse");
    vm.run(vybe_compiler_ruby::Compiler::new().compile(&rb_prog).expect("Ruby compile")).expect("Ruby run");

    assert_eq!(output.borrow().as_slice(), &["15"]);
}

#[test]
fn ruby_function_dart_calls() {
    let (mut vm, output) = setup_vm();
    let rb_src = "def square(x)\n  x * x\nend";
    let rb_prog = vybe_parser_ruby::parse(rb_src).expect("Ruby parse");
    vm.run(vybe_compiler_ruby::Compiler::new().compile(&rb_prog).expect("Ruby compile")).expect("Ruby run");

    let dart_src = "var r = square(9); print(r);";
    let dart_prog = vybe_parser_dart::parse(dart_src).expect("Dart parse");
    vm.run(vybe_compiler_dart::Compiler::new().compile(&dart_prog).expect("Dart compile")).expect("Dart run");

    assert_eq!(output.borrow().as_slice(), &["81"]);
}

// ═══════════════════════════════════════════════════════════
// SECTION 4: MULTI-LANGUAGE GLOBAL SHARING
// ═══════════════════════════════════════════════════════════

#[test]
fn five_languages_share_global() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);

    let vb = vybe_parser_basic::parse_program("Dim counter As Integer\ncounter = 1").expect("VB");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb).expect("VB")).expect("VB");

    let js = vybe_parser_js::parse("counter = counter + 10;").expect("JS");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js).expect("JS")).expect("JS");

    let rb = vybe_parser_ruby::parse("counter = counter + 100").expect("Ruby");
    vm.run(vybe_compiler_ruby::Compiler::new().compile(&rb).expect("Ruby")).expect("Ruby");

    let php = vybe_parser_php::parse("<?php $counter = $counter + 1000;").expect("PHP");
    vm.run(vybe_compiler_php::Compiler::new().compile(&php).expect("PHP")).expect("PHP");

    let py = vybe_parser_python::parse("print(counter)").expect("Python");
    vm.run(vybe_compiler_python::Compiler::new().compile(&py).expect("Python")).expect("Python");

    assert_eq!(output.borrow().as_slice(), &["1111"]);
}

#[test]
fn three_language_chain_compute() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);

    // VB sets value
    let vb = vybe_parser_basic::parse_program("Dim x As Integer\nx = 5").expect("VB");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb).expect("VB")).expect("VB");

    // JS squares it
    let js = vybe_parser_js::parse("x = x * x;").expect("JS");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js).expect("JS")).expect("JS");

    // C# prints
    let cs = vybe_parser_csharp::parse("Console.WriteLine(x);").expect("C#");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs).expect("C#")).expect("C#");

    assert_eq!(output.borrow().as_slice(), &["25"]);
}

// ═══════════════════════════════════════════════════════════
// SECTION 5: SAME-LANGUAGE CLASS TESTS (regression guard)
// ═══════════════════════════════════════════════════════════

#[test]
fn vb_class_with_inheritance() {
    let (mut vm, output) = setup_vm();
    let vb_src = r#"
Class Base
    Public Tag As String
    Public Sub New(t As String)
        Tag = t
    End Sub
End Class
Class Child
    Inherits Base
    Public Sub New()
        MyBase.New("child")
    End Sub
End Class
Dim c As New Child()
Console.WriteLine(c.Tag)
"#;
    let vb_prog = vybe_parser_basic::parse_program(vb_src).expect("VB parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&vb_prog).expect("VB compile")).expect("VB run");

    assert_eq!(output.borrow().as_slice(), &["child"]);
}

#[test]
fn js_class_with_inheritance() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);
    let js_src = r#"
class Base {
    constructor(tag) { this.tag = tag; }
}
class Child extends Base {
    constructor() { super("child"); }
}
var c = new Child();
console.log(c.tag);
"#;
    let js_prog = vybe_parser_js::parse(js_src).expect("JS parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&js_prog).expect("JS compile")).expect("JS run");

    assert_eq!(output.borrow().as_slice(), &["child"]);
}

#[test]
fn cs_class_with_inheritance() {
    let (mut vm, output) = setup_vm();
    let cs_src = r#"
class Base {
    public string Tag;
    public Base(string t) { Tag = t; }
}
class Child : Base {
    public Child() : base("child") {}
}
var c = new Child();
Console.WriteLine(c.Tag);
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile")).expect("C# run");

    assert_eq!(output.borrow().as_slice(), &["child"]);
}

// ═══════════════════════════════════════════════════════════
// SECTION 6: LINQ WITH HOST CALLBACKS
// ═══════════════════════════════════════════════════════════

#[test]
fn cs_linq_where_select_foreach() {
    let (mut vm, output) = setup_vm();
    let cs_src = r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3); list.Add(4); list.Add(5);
list.Where(x => x > 2).Select(x => x * 10).ForEach(x => Console.WriteLine(x));
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile")).expect("C# run");

    assert_eq!(output.borrow().as_slice(), &["30", "40", "50"]);
}

#[test]
fn cs_linq_any_all() {
    let (mut vm, output) = setup_vm();
    let cs_src = r#"
var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
Console.WriteLine(list.Any(x => x > 2));
Console.WriteLine(list.All(x => x > 0));
"#;
    let cs_prog = vybe_parser_csharp::parse(cs_src).expect("C# parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&cs_prog).expect("C# compile")).expect("C# run");

    assert_eq!(output.borrow().as_slice(), &["true", "true"]);
}

// ═══════════════════════════════════════════════════════════
// SECTION 7: COMPONENT MODEL ISOLATION
// ═══════════════════════════════════════════════════════════

#[test]
fn component_isolation_globals_separate() {
    let (mut vm, _) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);

    let js_chunks = vybe_compiler_js::Compiler::new()
        .compile(&vybe_parser_js::parse("var secret = 'js_data';").expect("JS")).expect("JS");
    let js_comp = vybe_compiler_common::components::build_component(
        "mod_js", vybe_bytecode::component::Language::JS, js_chunks);

    let rb_chunks = vybe_compiler_ruby::Compiler::new()
        .compile(&vybe_parser_ruby::parse("secret = 'ruby_data'").expect("Ruby")).expect("Ruby");
    let rb_comp = vybe_compiler_common::components::build_component(
        "mod_rb", vybe_bytecode::component::Language::Ruby, rb_chunks);

    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    linker.add_component(js_comp.clone());
    linker.add_component(rb_comp.clone());
    let link_result = linker.link().expect("Link failed");
    vm.run_components(&link_result, &[js_comp, rb_comp]).expect("run failed");

    assert!(vm.globals.contains_key("mod_js::secret"));
    assert!(vm.globals.contains_key("mod_rb::secret"));
    assert!(!vm.globals.contains_key("secret"));
}

// ═══════════════════════════════════════════════════════════
// SECTION 8: ALL 8 LANGUAGES PRODUCE OUTPUT
// ═══════════════════════════════════════════════════════════

#[test]
fn all_8_languages_output() {
    let (mut vm, output) = setup_vm();
    vybe_compiler_js::register_js_coercion(&mut vm);

    let tests: Vec<(&str, Box<dyn Fn(&mut VM)>)> = vec![
        ("VB", Box::new(|vm: &mut VM| {
            let p = vybe_parser_basic::parse_program("Console.WriteLine(\"VB\")").unwrap();
            vm.run(vybe_compiler_vb::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
        ("JS", Box::new(|vm: &mut VM| {
            let p = vybe_parser_js::parse("console.log('JS');").unwrap();
            vm.run(vybe_compiler_js::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
        ("CS", Box::new(|vm: &mut VM| {
            let p = vybe_parser_csharp::parse("Console.WriteLine(\"CS\");").unwrap();
            vm.run(vybe_compiler_csharp::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
        ("Ruby", Box::new(|vm: &mut VM| {
            let p = vybe_parser_ruby::parse("puts 'Ruby'").unwrap();
            vm.run(vybe_compiler_ruby::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
        ("PHP", Box::new(|vm: &mut VM| {
            let p = vybe_parser_php::parse("<?php echo 'PHP';").unwrap();
            vm.run(vybe_compiler_php::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
        ("Python", Box::new(|vm: &mut VM| {
            let p = vybe_parser_python::parse("print('Python')").unwrap();
            vm.run(vybe_compiler_python::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
        ("Dart", Box::new(|vm: &mut VM| {
            let p = vybe_parser_dart::parse("print('Dart');").unwrap();
            vm.run(vybe_compiler_dart::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
        ("COBOL", Box::new(|vm: &mut VM| {
            let p = vybe_parser_cobol::parse("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    DISPLAY \"COBOL\".\n    STOP RUN.").unwrap();
            vm.run(vybe_compiler_cobol::Compiler::new().compile(&p).unwrap()).unwrap();
        })),
    ];

    for (_, run_fn) in &tests {
        run_fn(&mut vm);
    }

    let out = output.borrow();
    assert_eq!(out.len(), 8);
    assert_eq!(&out[..], &["VB", "JS", "CS", "Ruby", "PHP", "Python", "Dart", "COBOL"]);
}
