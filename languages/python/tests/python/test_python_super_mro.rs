// Python super() and MRO — cooperative multiple inheritance, super() calls
use super::helpers::run_python;

#[test]
fn test_super_basic() {
    let script = r#"
class Animal:
    def speak(self):
        return "..."

class Dog(Animal):
    def speak(self):
        parent = super().speak()
        return f"Woof! (parent: {parent})"

d = Dog()
print(d.speak())
"#;
    assert_eq!(run_python(script), vec!["Woof! (parent: ...)"]);
}

#[test]
fn test_super_in_init() {
    let script = r#"
class Base:
    def __init__(self, x):
        self.x = x

class Child(Base):
    def __init__(self, x, y):
        super().__init__(x)
        self.y = y

c = Child(10, 20)
print(c.x, c.y)
"#;
    assert_eq!(run_python(script), vec!["10 20"]);
}

#[test]
fn test_mro_diamond() {
    let script = r#"
class A:
    def who(self):
        return "A"

class B(A):
    def who(self):
        return "B->" + super().who()

class C(A):
    def who(self):
        return "C->" + super().who()

class D(B, C):
    def who(self):
        return "D->" + super().who()

d = D()
print(d.who())
print([cls.__name__ for cls in D.__mro__])
"#;
    assert_eq!(run_python(script), vec!["D->B->C->A", "['D', 'B', 'C', 'A', 'object']"]);
}

#[test]
fn test_super_skips_current_class() {
    let script = r#"
class A:
    val = "A"

class B(A):
    val = "B"

class C(B):
    def get_val(self):
        return super().val  # skips C, gets B

print(C().get_val())
"#;
    assert_eq!(run_python(script), vec!["B"]);
}

#[test]
fn test_cooperative_init_chain() {
    let script = r#"
class Mixin1:
    def __init__(self, **kwargs):
        print("Mixin1")
        super().__init__(**kwargs)

class Mixin2:
    def __init__(self, **kwargs):
        print("Mixin2")
        super().__init__(**kwargs)

class Base:
    def __init__(self, **kwargs):
        print("Base")

class MyClass(Mixin1, Mixin2, Base):
    def __init__(self):
        super().__init__()

MyClass()
"#;
    assert_eq!(run_python(script), vec!["Mixin1", "Mixin2", "Base"]);
}

#[test]
fn test_super_classmethod() {
    let script = r#"
class Counter:
    count = 0

    @classmethod
    def increment(cls):
        cls.count += 1

class DoubleCounter(Counter):
    @classmethod
    def increment(cls):
        super().increment()
        super().increment()

DoubleCounter.increment()
print(DoubleCounter.count)
"#;
    assert_eq!(run_python(script), vec!["2"]);
}
