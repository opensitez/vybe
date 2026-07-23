use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Class Inheritance & MRO — C3 linearization, super(), method overriding, attribute resolution order
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_single_inheritance_method_override() {
    let src = r#"
class Base:
    def greet(self):
        return "Base greeting"

class Child(Base):
    def greet(self):
        return "Child greeting"

print(Base().greet())
print(Child().greet())
"#;
    assert_eq!(run_python(src), vec!["Base greeting", "Child greeting"]);
}

#[test]
fn test_py_super_explicit_arguments() {
    let src = r#"
class Parent:
    def __init__(self, name):
        self.name = name

class Child(Parent):
    def __init__(self, name, age):
        super(Child, self).__init__(name)
        self.age = age

c = Child("Alice", 25)
print(c.name, c.age)
"#;
    assert_eq!(run_python(src), vec!["Alice 25"]);
}

#[test]
fn test_py_zero_argument_super() {
    let src = r#"
class Parent:
    def describe(self):
        return "Parent"

class Child(Parent):
    def describe(self):
        return super().describe() + " -> Child"

print(Child().describe())
"#;
    assert_eq!(run_python(src), vec!["Parent -> Child"]);
}

#[test]
fn test_py_diamond_inheritance_mro() {
    let src = r#"
class A:
    def who(self): return "A"

class B(A):
    def who(self): return "B->" + super().who()

class C(A):
    def who(self): return "C->" + super().who()

class D(B, C):
    def who(self): return "D->" + super().who()

d = D()
print(d.who())
mro_names = [cls.__name__ for cls in D.__mro__]
print(mro_names)
"#;
    assert_eq!(
        run_python(src),
        vec!["D->B->C->A", "['D', 'B', 'C', 'A', 'object']"]
    );
}

#[test]
fn test_py_invalid_mro_c3_linearization_error() {
    let src = r#"
class X: pass
class Y: pass

try:
    # A(X, Y) and B(Y, X) creates inconsistent hierarchy C(A, B)
    class A(X, Y): pass
    class B(Y, X): pass
    class C(A, B): pass
except TypeError as e:
    print("TypeError: Cannot create a consistent method resolution order")
"#;
    assert_eq!(
        run_python(src),
        vec!["TypeError: Cannot create a consistent method resolution order"]
    );
}

#[test]
fn test_py_issubclass_and_isinstance_inheritance() {
    let src = r#"
class A: pass
class B(A): pass
class C(B): pass

c = C()
print(isinstance(c, C))
print(isinstance(c, B))
print(isinstance(c, A))
print(issubclass(C, A))
print(issubclass(B, C))
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "True", "True", "False"]
    );
}

#[test]
fn test_py_mixin_pattern_multiple_inheritance() {
    let src = r#"
class JSONMixin:
    def to_json(self):
        import json
        return json.dumps(self.__dict__)

class Model:
    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)

class User(Model, JSONMixin):
    pass

u = User(name="Bob", age=30)
print(u.to_json())
"#;
    assert_eq!(run_python(src), vec!["{\"name\": \"Bob\", \"age\": 30}"]);
}

#[test]
fn test_py_class_attribute_shadowing_by_instance() {
    let src = r#"
class Parent:
    species = "Homo Sapiens"

p1 = Parent()
p2 = Parent()

print(p1.species, p2.species)
p1.species = "Mutant"  # instance shadow
print(p1.species, p2.species, Parent.species)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Homo Sapiens Homo Sapiens",
            "Mutant Homo Sapiens Homo Sapiens"
        ]
    );
}

#[test]
fn test_py_super_in_classmethod() {
    let src = r#"
class Base:
    @classmethod
    def create(cls):
        return f"Base create for {cls.__name__}"

class Child(Base):
    @classmethod
    def create(cls):
        return super().create() + " extended"

print(Child.create())
"#;
    assert_eq!(run_python(src), vec!["Base create for Child extended"]);
}

#[test]
fn test_py_private_attribute_mangling_in_inheritance() {
    let src = r#"
class Base:
    def __init__(self):
        self.__secret = 42

    def get_secret(self):
        return self.__secret

class Child(Base):
    def __init__(self):
        super().__init__()
        self.__secret = 99  # mangled to _Child__secret, doesn't overwrite Base

c = Child()
print(c.get_secret())
print(hasattr(c, "_Base__secret"))
print(hasattr(c, "_Child__secret"))
"#;
    assert_eq!(run_python(src), vec!["42", "True", "True"]);
}
