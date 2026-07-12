use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// Inheritance runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn inheritance_method_override() {
    let out = run_python(
        r#"
class Animal:
    def speak(self):
        return "..."
class Dog(Animal):
    def speak(self):
        return "Woof"
d = Dog()
print(d.speak())
"#,
    );
    assert_eq!(out[0], "Woof");
}

#[test]
fn super_call() {
    compile_ok(
        "class Dog(Animal):\n    def __init__(self, name):\n        super().__init__(name)\n",
    );
}

#[test]
fn super_method_call() {
    compile_ok(
        "class Child(Parent):\n    def method(self):\n        return super().method() + 1\n",
    );
}

// ══════════════════════════════════════════════════════════════════════════════
// Class variables vs instance variables
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn class_variable_access() {
    let out = run_python(
        r#"
class Config:
    version = "1.0"
print(Config.version)
"#,
    );
    assert_eq!(out[0], "1.0");
}

// ══════════════════════════════════════════════════════════════════════════════
// Dunder methods
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn dunder_len() {
    compile_ok(
        "class Bag:\n    def __init__(self):\n        self.items = []\n    def __len__(self):\n        return len(self.items)\n",
    );
}

#[test]
fn dunder_repr() {
    compile_ok(
        "class Point:\n    def __init__(self, x, y):\n        self.x = x\n        self.y = y\n    def __repr__(self):\n        return f'Point({self.x}, {self.y})'\n",
    );
}

#[test]
fn dunder_add() {
    compile_ok(
        "class Vec:\n    def __init__(self, x, y):\n        self.x = x\n        self.y = y\n    def __add__(self, other):\n        return Vec(self.x + other.x, self.y + other.y)\n",
    );
}

#[test]
fn dunder_eq() {
    compile_ok(
        "class Point:\n    def __init__(self, x, y):\n        self.x = x\n        self.y = y\n    def __eq__(self, other):\n        return self.x == other.x and self.y == other.y\n",
    );
}

#[test]
fn dunder_getitem() {
    compile_ok(
        "class Row:\n    def __init__(self, data):\n        self.data = data\n    def __getitem__(self, key):\n        return self.data[key]\n",
    );
}

// ══════════════════════════════════════════════════════════════════════════════
// Static methods and properties
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn staticmethod_compile() {
    compile_ok("class Math:\n    @staticmethod\n    def add(a, b):\n        return a + b\n");
}

#[test]
fn classmethod_compile() {
    compile_ok(
        "class Factory:\n    count = 0\n    @classmethod\n    def create(cls):\n        cls.count += 1\n        return Factory()\n",
    );
}

#[test]
fn property_compile() {
    compile_ok(
        "class Circle:\n    def __init__(self, r):\n        self._r = r\n    @property\n    def radius(self):\n        return self._r\n",
    );
}

// ══════════════════════════════════════════════════════════════════════════════
// Complex class patterns
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn class_with_metaclass() {
    compile_ok("class Meta(type):\n    pass\nclass MyClass(metaclass=Meta):\n    pass\n");
}

#[test]
fn class_multiple_bases() {
    compile_ok("class A:\n    pass\nclass B:\n    pass\nclass C(A, B):\n    pass\n");
}

#[test]
fn class_nested() {
    compile_ok("class Outer:\n    class Inner:\n        pass\n");
}

#[test]
fn class_method_calls_other() {
    let out = run_python(
        r#"
class Calculator:
    def __init__(self):
        self.result = 0
    def add(self, n):
        self.result = self.result + n
        return self
    def get(self):
        return self.result
c = Calculator()
c.add(5)
c.add(3)
print(c.get())
"#,
    );
    assert_eq!(out[0], "8");
}
