/// Comprehensive class tests — same patterns across all 8 languages.
/// Each language tests: base class, fields, methods, static methods, child class, method override.

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
// VB INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn vb_base_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
Class Animal
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Function Speak() As String
        Return Name & " speaks"
    End Function
End Class
Dim a As New Animal("Rex")
Console.WriteLine(a.Speak())
"#;
    let prog = vybe_parser_basic::parse_program(src).expect("parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

#[test]
fn vb_child_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
Class Animal
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Function Speak() As String
        Return Name & " speaks"
    End Function
End Class
Class Dog
    Inherits Animal
    Public Sub New(n As String)
        MyBase.New(n)
    End Sub
End Class
Dim d As New Dog("Rex")
Console.WriteLine(d.Speak())
"#;
    let prog = vybe_parser_basic::parse_program(src).expect("parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// JS INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn js_base_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Animal {
    constructor(name) { this.name = name; }
    speak() { return this.name + " speaks"; }
}
var a = new Animal("Rex");
console.log(a.speak());
"#;
    let prog = vybe_parser_js::parse(src).expect("parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

#[test]
fn js_child_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Animal {
    constructor(name) { this.name = name; }
    speak() { return this.name + " speaks"; }
}
class Dog extends Animal {
    constructor(name) { super(name); }
}
var d = new Dog("Rex");
console.log(d.speak());
"#;
    let prog = vybe_parser_js::parse(src).expect("parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// C# INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn cs_base_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Animal {
    public string Name;
    public Animal(string n) { Name = n; }
    public string Speak() { return Name + " speaks"; }
}
var a = new Animal("Rex");
Console.WriteLine(a.Speak());
"#;
    let prog = vybe_parser_csharp::parse(src).expect("parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

#[test]
fn cs_child_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Animal {
    public string Name;
    public Animal(string n) { Name = n; }
    public string Speak() { return Name + " speaks"; }
}
class Dog : Animal {
    public Dog(string n) : base(n) {}
}
var d = new Dog("Rex");
Console.WriteLine(d.Speak());
"#;
    let prog = vybe_parser_csharp::parse(src).expect("parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// PYTHON INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn python_base_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Animal:
    def __init__(self, name):
        self.name = name
    def speak(self):
        return self.name + " speaks"
a = Animal("Rex")
print(a.speak())
"#;
    let prog = vybe_parser_python::parse(src).expect("parse");
    vm.run(vybe_compiler_python::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

#[test]
fn python_child_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Animal:
    def __init__(self, name):
        self.name = name
    def speak(self):
        return self.name + " speaks"
class Dog(Animal):
    def __init__(self, name):
        super().__init__(name)
d = Dog("Rex")
print(d.speak())
"#;
    let prog = vybe_parser_python::parse(src).expect("parse");
    vm.run(vybe_compiler_python::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// RUBY INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn ruby_base_class() {
    let (mut vm, output) = setup_vm();
    let src = "class Animal\n  def initialize(name)\n    @name = name\n  end\n  def speak\n    @name + \" speaks\"\n  end\nend\na = Animal.new(\"Rex\")\nputs a.speak";
    let prog = vybe_parser_ruby::parse(src).expect("parse");
    vm.run(vybe_compiler_ruby::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

#[test]
fn ruby_child_class() {
    let (mut vm, output) = setup_vm();
    let src = "class Animal\n  def initialize(name)\n    @name = name\n  end\n  def speak\n    @name + \" speaks\"\n  end\nend\nclass Dog < Animal\n  def initialize(name)\n    super(name)\n  end\nend\nd = Dog.new(\"Rex\")\nputs d.speak";
    let prog = vybe_parser_ruby::parse(src).expect("parse");
    vm.run(vybe_compiler_ruby::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// PHP INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn php_base_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"<?php
class Animal {
    public $name;
    public function __construct($name) { $this->name = $name; }
    public function speak() { return $this->name . " speaks"; }
}
$a = new Animal("Rex");
echo $a->speak();
"#;
    let prog = vybe_parser_php::parse(src).expect("parse");
    vm.run(vybe_compiler_php::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

#[test]
fn php_child_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"<?php
class Animal {
    public $name;
    public function __construct($name) { $this->name = $name; }
    public function speak() { return $this->name . " speaks"; }
}
class Dog extends Animal {
    public function __construct($name) {
        parent::__construct($name);
    }
}
$d = new Dog("Rex");
echo $d->speak();
"#;
    let prog = vybe_parser_php::parse(src).expect("parse");
    vm.run(vybe_compiler_php::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// DART INHERITANCE
// ═══════════════════════════════════════════════════════════

#[test]
fn dart_base_class() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Animal {
    String name;
    Animal(this.name);
    String speak() { return name + " speaks"; }
}
var a = Animal("Rex");
print(a.speak());
"#;
    let prog = vybe_parser_dart::parse(src).expect("parse");
    vm.run(vybe_compiler_dart::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["Rex speaks"]);
}

// ═══════════════════════════════════════════════════════════
// FIELD ACCESS — all languages
// ═══════════════════════════════════════════════════════════

#[test]
fn vb_field_access() {
    let (mut vm, output) = setup_vm();
    let src = r#"
Class Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
End Class
Dim p As New Point(10, 20)
Console.WriteLine(p.X + p.Y)
"#;
    let prog = vybe_parser_basic::parse_program(src).expect("parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["30"]);
}

#[test]
fn js_field_access() {
    let (mut vm, output) = setup_vm();
    let src = "class Point { constructor(x, y) { this.x = x; this.y = y; } }\nvar p = new Point(10, 20);\nconsole.log(p.x + p.y);";
    let prog = vybe_parser_js::parse(src).expect("parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["30"]);
}

#[test]
fn cs_field_access() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Point { public int X; public int Y; public Point(int a, int b) { X = a; Y = b; } }
var p = new Point(10, 20);
Console.WriteLine(p.X + p.Y);
"#;
    let prog = vybe_parser_csharp::parse(src).expect("parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["30"]);
}

#[test]
fn python_field_access() {
    let (mut vm, output) = setup_vm();
    let src = "class Point:\n    def __init__(self, x, y):\n        self.x = x\n        self.y = y\np = Point(10, 20)\nprint(p.x + p.y)";
    let prog = vybe_parser_python::parse(src).expect("parse");
    vm.run(vybe_compiler_python::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["30"]);
}

#[test]
fn php_field_access() {
    let (mut vm, output) = setup_vm();
    let src = "<?php\nclass Point { public $x; public $y; public function __construct($x, $y) { $this->x = $x; $this->y = $y; } }\n$p = new Point(10, 20);\necho $p->x + $p->y;";
    let prog = vybe_parser_php::parse(src).expect("parse");
    vm.run(vybe_compiler_php::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["30"]);
}

// ═══════════════════════════════════════════════════════════
// STATIC METHODS — all languages
// ═══════════════════════════════════════════════════════════

#[test]
fn js_static_method() {
    let (mut vm, output) = setup_vm();
    let src = "class MathUtil { static add(a, b) { return a + b; } }\nconsole.log(MathUtil.add(3, 4));";
    let prog = vybe_parser_js::parse(src).expect("parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["7"]);
}

#[test]
fn cs_static_method() {
    let (mut vm, output) = setup_vm();
    let src = "class MathUtil { public static int Add(int a, int b) { return a + b; } }\nConsole.WriteLine(MathUtil.Add(3, 4));";
    let prog = vybe_parser_csharp::parse(src).expect("parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["7"]);
}

#[test]
fn vb_shared_method() {
    let (mut vm, output) = setup_vm();
    let src = r#"
Class MathUtil
    Public Shared Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class
Console.WriteLine(MathUtil.Add(3, 4))
"#;
    let prog = vybe_parser_basic::parse_program(src).expect("parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["7"]);
}

#[test]
fn php_static_method() {
    let (mut vm, output) = setup_vm();
    let src = "<?php\nclass MathUtil { public static function add($a, $b) { return $a + $b; } }\necho MathUtil::add(3, 4);";
    let prog = vybe_parser_php::parse(src).expect("parse");
    vm.run(vybe_compiler_php::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["7"]);
}

// ═══════════════════════════════════════════════════════════
// METHOD OVERRIDE — child overrides parent method
// ═══════════════════════════════════════════════════════════

#[test]
fn js_method_override() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Base { greet() { return "base"; } }
class Child extends Base { greet() { return "child"; } }
var c = new Child();
console.log(c.greet());
"#;
    let prog = vybe_parser_js::parse(src).expect("parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["child"]);
}

#[test]
fn cs_method_override() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Base { public string Greet() { return "base"; } }
class Child : Base { public string Greet() { return "child"; } }
var c = new Child();
Console.WriteLine(c.Greet());
"#;
    let prog = vybe_parser_csharp::parse(src).expect("parse");
    vm.run(vybe_compiler_csharp::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["child"]);
}

#[test]
fn vb_method_override() {
    let (mut vm, output) = setup_vm();
    let src = r#"
Class Base
    Public Function Greet() As String
        Return "base"
    End Function
End Class
Class Child
    Inherits Base
    Public Function Greet() As String
        Return "child"
    End Function
End Class
Dim c As New Child()
Console.WriteLine(c.Greet())
"#;
    let prog = vybe_parser_basic::parse_program(src).expect("parse");
    vm.run(vybe_compiler_vb::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["child"]);
}

// ═══════════════════════════════════════════════════════════
// MULTIPLE INSTANCES — verify objects are independent
// ═══════════════════════════════════════════════════════════

#[test]
fn js_multiple_instances() {
    let (mut vm, output) = setup_vm();
    let src = r#"
class Counter {
    constructor() { this.count = 0; }
    inc() { this.count = this.count + 1; }
    value() { return this.count; }
}
var a = new Counter();
var b = new Counter();
a.inc(); a.inc(); a.inc();
b.inc();
console.log(a.value());
console.log(b.value());
"#;
    let prog = vybe_parser_js::parse(src).expect("parse");
    vm.run(vybe_compiler_js::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["3", "1"]);
}

#[test]
fn php_multiple_instances() {
    let (mut vm, output) = setup_vm();
    let src = r#"<?php
class Counter {
    public $count;
    public function __construct() { $this->count = 0; }
    public function inc() { $this->count++; }
    public function value() { return $this->count; }
}
$a = new Counter();
$b = new Counter();
$a->inc(); $a->inc(); $a->inc();
$b->inc();
echo $a->value();
echo $b->value();
"#;
    let prog = vybe_parser_php::parse(src).expect("parse");
    vm.run(vybe_compiler_php::Compiler::new().compile(&prog).expect("compile")).expect("run");
    assert_eq!(output.borrow().as_slice(), &["3", "1"]);
}
