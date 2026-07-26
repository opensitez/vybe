// Python property advanced — computed, cached, deleter, class property patterns
use super::helpers::run_python;

#[test]
fn test_property_getter() {
    let script = r#"
class Circle:
    def __init__(self, radius):
        self._radius = radius

    @property
    def radius(self):
        return self._radius

    @property
    def area(self):
        import math
        return math.pi * self._radius ** 2

c = Circle(5)
print(c.radius)
print(round(c.area, 4))
"#;
    assert_eq!(run_python(script), vec!["5", "78.5398"]);
}

#[test]
fn test_property_setter() {
    let script = r#"
class Temperature:
    def __init__(self, celsius):
        self._celsius = celsius

    @property
    def celsius(self):
        return self._celsius

    @celsius.setter
    def celsius(self, value):
        if value < -273.15:
            raise ValueError("Below absolute zero")
        self._celsius = value

    @property
    def fahrenheit(self):
        return self._celsius * 9 / 5 + 32

t = Temperature(100)
print(t.fahrenheit)
t.celsius = 0
print(t.fahrenheit)
"#;
    assert_eq!(run_python(script), vec!["212.0", "32.0"]);
}

#[test]
fn test_property_deleter() {
    let script = r#"
class DataHolder:
    def __init__(self, data):
        self._data = data

    @property
    def data(self):
        if self._data is None:
            raise AttributeError("data deleted")
        return self._data

    @data.deleter
    def data(self):
        self._data = None

d = DataHolder([1, 2, 3])
print(d.data)
del d.data
try:
    _ = d.data
    print("still there")
except AttributeError:
    print("deleted")
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3]", "deleted"]);
}

#[test]
fn test_property_validation() {
    let script = r#"
class BoundedInt:
    def __init__(self, value):
        self.value = value  # uses setter

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, v):
        self._value = max(0, min(100, v))

b = BoundedInt(50)
print(b.value)
b.value = 150
print(b.value)
b.value = -10
print(b.value)
"#;
    assert_eq!(run_python(script), vec!["50", "100", "0"]);
}

#[test]
fn test_property_inheritance() {
    let script = r#"
class Base:
    @property
    def name(self):
        return "Base"

class Child(Base):
    @property
    def name(self):
        return "Child"

b = Base()
c = Child()
print(b.name)
print(c.name)
"#;
    assert_eq!(run_python(script), vec!["Base", "Child"]);
}
