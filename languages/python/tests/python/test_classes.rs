use super::helpers::{compile_ok, run_python};

#[test]
fn class_basic() {
    compile_ok("class Foo:\n    def bar(self):\n        print('bar')\n");
}

#[test]
fn class_with_init() {
    compile_ok(
        "class Dog:\n    def __init__(self, name):\n        self.name = name\n    def bark(self):\n        print(self.name)\n",
    );
}

#[test]
fn class_with_inheritance() {
    compile_ok(
        "class Animal:\n    def speak(self):\n        pass\n\nclass Dog(Animal):\n    def bark(self):\n        pass\n",
    );
}

#[test]
fn multiple_classes() {
    compile_ok(
        "class Cat:\n    def meow(self):\n        pass\n\nclass Dog:\n    def bark(self):\n        pass\n",
    );
}

#[test]
fn no_args_class() {
    compile_ok("class Empty:\n    pass\n");
}

#[test]
fn class_with_only_methods() {
    compile_ok(
        r#"
class Calculator:
    def add(self, a, b):
        return a + b
    def sub(self, a, b):
        return a - b
"#,
    );
}

// Multiple inheritance

#[test]
fn single_inheritance() {
    compile_ok(
        "class Animal:\n    def speak(self):\n        return 'generic'\n\nclass Dog(Animal):\n    def speak(self):\n        return 'woof'\n",
    );
}

#[test]
fn multiple_inheritance() {
    compile_ok(
        "class A:\n    def method_a(self):\n        return 'a'\n\nclass B:\n    def method_b(self):\n        return 'b'\n\nclass C(A, B):\n    pass\n",
    );
}

#[test]
fn diamond_inheritance() {
    compile_ok(
        "class Base:\n    pass\nclass Left(Base):\n    pass\nclass Right(Base):\n    pass\nclass Child(Left, Right):\n    pass\n",
    );
}

// @staticmethod / @classmethod

#[test]
fn staticmethod_basic() {
    compile_ok(
        "class Math:\n    @staticmethod\n    def add(a, b):\n        return a + b\nresult = Math.add(1, 2)\n",
    );
}

#[test]
fn staticmethod_no_self_param() {
    compile_ok(
        "class Config:\n    @staticmethod\n    def default_value():\n        return 42\nv = Config.default_value()\n",
    );
}

#[test]
fn classmethod_basic() {
    compile_ok(
        "class Foo:\n    @classmethod\n    def create(cls):\n        return Foo()\n    def __init__(self):\n        pass\n",
    );
}

// @property

#[test]
fn class_property_decorator() {
    compile_ok(
        r#"
class Circle:
    def __init__(self, r):
        self.r = r
    @property
    def radius(self):
        return self.r
"#,
    );
}

// Runtime class tests

#[test]
fn class_instance_has_properties() {
    let out = run_python(
        r#"
class Dog:
    def __init__(self, name, age):
        self.name = name
        self.age = age
d = Dog("Rex", 3)
print(d.name)
print(d.age)
"#,
    );
    assert_eq!(out[0], "Rex");
    assert_eq!(out[1], "3");
}

#[test]
fn class_method_returns_value() {
    let out = run_python(
        r#"
class Dog:
    def __init__(self, name):
        self.name = name
    def bark(self):
        return self.name
d = Dog("Rex")
print(d.bark())
"#,
    );
    assert_eq!(out[0], "Rex");
}

#[test]
fn class_multiple_instances() {
    let out = run_python(
        r#"
class Counter:
    def __init__(self):
        self.count = 0
    def inc(self):
        self.count = self.count + 1
    def get(self):
        return self.count
a = Counter()
b = Counter()
a.inc()
a.inc()
b.inc()
print(a.get())
print(b.get())
"#,
    );
    assert_eq!(out[0], "2");
    assert_eq!(out[1], "1");
}

#[test]
fn class_method_self_access() {
    let out = run_python(
        r#"
class Person:
    def __init__(self, first, last):
        self.first = first
        self.last = last
    def full_name(self):
        return self.first + " " + self.last
p = Person("John", "Doe")
print(p.full_name())
"#,
    );
    assert_eq!(out[0], "John Doe");
}

#[test]
fn class_method_modifies_state() {
    let out = run_python(
        r#"
class Stack:
    def __init__(self):
        self.items = []
    def push(self, item):
        self.items.append(item)
    def size(self):
        return len(self.items)
s = Stack()
s.push(1)
s.push(2)
s.push(3)
print(s.size())
"#,
    );
    assert_eq!(out[0], "3");
}

// User methods override builtins

#[test]
fn user_method_named_get() {
    compile_ok("class C:\n    def get(self):\n        return 42\nc = C()\nprint(c.get())\n");
}

#[test]
fn user_method_named_append() {
    compile_ok("class C:\n    def append(self, x):\n        print(x)\nc = C()\nc.append(42)\n");
}

// Enum-like class

#[test]
fn enum_class() {
    compile_ok("class Color:\n    RED = 1\n    GREEN = 2\n    BLUE = 3\nc = Color()\n");
}
