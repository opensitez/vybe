// Python abc module — ABCMeta, abstractmethod, ABC, abstractproperty, register
use super::helpers::run_python;

#[test]
fn test_abc_abstractmethod_prevents_instantiation() {
    let script = r#"
from abc import ABC, abstractmethod

class Shape(ABC):
    @abstractmethod
    def area(self):
        pass

try:
    s = Shape()
    print("no_error")
except TypeError:
    print("TypeError")
"#;
    assert_eq!(run_python(script), vec!["TypeError"]);
}

#[test]
fn test_abc_concrete_subclass_instantiates() {
    let script = r#"
from abc import ABC, abstractmethod

class Animal(ABC):
    @abstractmethod
    def speak(self):
        pass

class Dog(Animal):
    def speak(self):
        return "woof"

d = Dog()
print(d.speak())
"#;
    assert_eq!(run_python(script), vec!["woof"]);
}

#[test]
fn test_abc_multiple_abstract_methods() {
    let script = r#"
from abc import ABC, abstractmethod

class Vehicle(ABC):
    @abstractmethod
    def fuel_type(self):
        pass
    @abstractmethod
    def max_speed(self):
        pass

class Car(Vehicle):
    def fuel_type(self):
        return "petrol"
    def max_speed(self):
        return 200

c = Car()
print(c.fuel_type())
print(c.max_speed())
"#;
    assert_eq!(run_python(script), vec!["petrol", "200"]);
}

#[test]
fn test_abc_register_virtual_subclass() {
    let script = r#"
from abc import ABC

class Printable(ABC):
    pass

class Doc:
    pass

Printable.register(Doc)
print(issubclass(Doc, Printable))
print(isinstance(Doc(), Printable))
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_abc_isabstract_check() {
    let script = r#"
import abc
from abc import ABC, abstractmethod

class Base(ABC):
    @abstractmethod
    def do(self):
        pass

print(abc.isabstract(Base))

class Concrete(Base):
    def do(self):
        return 1

print(abc.isabstract(Concrete))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_abc_abstract_property() {
    let script = r#"
from abc import ABC, abstractmethod

class Config(ABC):
    @property
    @abstractmethod
    def name(self):
        pass

class AppConfig(Config):
    @property
    def name(self):
        return "MyApp"

cfg = AppConfig()
print(cfg.name)
"#;
    assert_eq!(run_python(script), vec!["MyApp"]);
}

#[test]
fn test_abc_partial_implementation_still_abstract() {
    let script = r#"
from abc import ABC, abstractmethod

class Base(ABC):
    @abstractmethod
    def a(self):
        pass
    @abstractmethod
    def b(self):
        pass

class Partial(Base):
    def a(self):
        return 1

try:
    Partial()
    print("no_error")
except TypeError:
    print("TypeError_still_abstract")
"#;
    assert_eq!(run_python(script), vec!["TypeError_still_abstract"]);
}

#[test]
fn test_abc_mro_respected() {
    let script = r#"
from abc import ABC, abstractmethod

class A(ABC):
    @abstractmethod
    def method(self):
        pass

class B(A):
    def method(self):
        return "B"

class C(B):
    pass

c = C()
print(c.method())
print(isinstance(c, A))
"#;
    assert_eq!(run_python(script), vec!["B", "True"]);
}
