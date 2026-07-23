use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: class OOP — inheritance, MRO, super(), dunder methods, slots, properties, classmethods, staticmethods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_class_basic_inheritance() {
    let src = r#"
class Animal:
    def __init__(self, name):
        self.name = name

    def speak(self):
        return f"{self.name} makes a sound"

class Dog(Animal):
    def speak(self):
        return f"{self.name} says Woof!"

class Cat(Animal):
    def speak(self):
        return f"{self.name} says Meow!"

animals = [Dog("Rex"), Cat("Whiskers"), Animal("Unknown")]
for a in animals:
    print(a.speak())
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Rex says Woof!",
            "Whiskers says Meow!",
            "Unknown makes a sound"
        ]
    );
}

#[test]
fn test_py_class_super_cooperative_multiple_inheritance() {
    let src = r#"
class A:
    def greet(self):
        return "A"

class B(A):
    def greet(self):
        return "B->" + super().greet()

class C(A):
    def greet(self):
        return "C->" + super().greet()

class D(B, C):
    def greet(self):
        return "D->" + super().greet()

print(D().greet())
print(D.__mro__)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "D->B->C->A",
            "(<class '__main__.D'>, <class '__main__.B'>, <class '__main__.C'>, <class '__main__.A'>, <class 'object'>)"
        ]
    );
}

#[test]
fn test_py_class_dunder_repr_str() {
    let src = r#"
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __repr__(self):
        return f"Point({self.x!r}, {self.y!r})"

    def __str__(self):
        return f"({self.x}, {self.y})"

p = Point(3, 4)
print(repr(p))
print(str(p))
print(f"{p}")   # uses __str__
print(f"{p!r}") # uses __repr__
"#;
    assert_eq!(
        run_python(src),
        vec!["Point(3, 4)", "(3, 4)", "(3, 4)", "Point(3, 4)"]
    );
}

#[test]
fn test_py_class_dunder_len_getitem_contains() {
    let src = r#"
class Bag:
    def __init__(self, *items):
        self._items = list(items)

    def __len__(self):
        return len(self._items)

    def __getitem__(self, idx):
        return self._items[idx]

    def __contains__(self, item):
        return item in self._items

b = Bag("a", "b", "c")
print(len(b))
print(b[1])
print("b" in b)
print(list(b))  # iteration via __getitem__
"#;
    assert_eq!(run_python(src), vec!["3", "b", "True", "['a', 'b', 'c']"]);
}

#[test]
fn test_py_class_dunder_arithmetic_operators() {
    let src = r#"
class Vector:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __add__(self, other):
        return Vector(self.x + other.x, self.y + other.y)

    def __mul__(self, scalar):
        return Vector(self.x * scalar, self.y * scalar)

    def __rmul__(self, scalar):
        return self.__mul__(scalar)

    def __abs__(self):
        return (self.x ** 2 + self.y ** 2) ** 0.5

    def __repr__(self):
        return f"Vector({self.x}, {self.y})"

v1 = Vector(1, 2)
v2 = Vector(3, 4)
print(v1 + v2)
print(v1 * 3)
print(3 * v1)
print(round(abs(v2), 2))
"#;
    assert_eq!(
        run_python(src),
        vec!["Vector(4, 6)", "Vector(3, 6)", "Vector(3, 6)", "5.0"]
    );
}

#[test]
fn test_py_class_comparison_operators() {
    let src = r#"
class Money:
    def __init__(self, amount):
        self.amount = amount

    def __eq__(self, other):
        return self.amount == other.amount

    def __lt__(self, other):
        return self.amount < other.amount

    def __le__(self, other):
        return self.amount <= other.amount

    def __repr__(self):
        return f"Money({self.amount})"

m1 = Money(10)
m2 = Money(20)
m3 = Money(10)
print(m1 < m2)
print(m1 == m3)
print(m2 > m1)
print(sorted([m2, m1, m3]))
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "True", "[Money(10), Money(10), Money(20)]"]
    );
}

#[test]
fn test_py_class_slots() {
    let src = r#"
class Compact:
    __slots__ = ('x', 'y', 'z')

    def __init__(self, x, y, z):
        self.x = x
        self.y = y
        self.z = z

c = Compact(1, 2, 3)
print(c.x, c.y, c.z)
print(hasattr(c, '__dict__'))
try:
    c.new_attr = "extra"
except AttributeError:
    print("AttributeError: no __dict__")
"#;
    assert_eq!(
        run_python(src),
        vec!["1 2 3", "False", "AttributeError: no __dict__"]
    );
}

#[test]
fn test_py_class_property_getter_setter_deleter() {
    let src = r#"
class Temperature:
    def __init__(self):
        self._celsius = 0

    @property
    def celsius(self):
        return self._celsius

    @celsius.setter
    def celsius(self, val):
        self._celsius = val

    @celsius.deleter
    def celsius(self):
        print("Deleting temperature")
        del self._celsius

    @property
    def fahrenheit(self):
        return self._celsius * 9/5 + 32

t = Temperature()
t.celsius = 100
print(t.celsius)
print(t.fahrenheit)
del t.celsius
"#;
    assert_eq!(
        run_python(src),
        vec!["100", "212.0", "Deleting temperature"]
    );
}

#[test]
fn test_py_class_classmethod_and_factory() {
    let src = r#"
class Date:
    def __init__(self, year, month, day):
        self.year = year
        self.month = month
        self.day = day

    @classmethod
    def from_string(cls, date_str):
        year, month, day = map(int, date_str.split('-'))
        return cls(year, month, day)

    @classmethod
    def today_placeholder(cls):
        return cls(2024, 1, 1)

    def __repr__(self):
        return f"Date({self.year}, {self.month}, {self.day})"

d = Date.from_string("2024-06-15")
print(d)
print(Date.today_placeholder())
"#;
    assert_eq!(
        run_python(src),
        vec!["Date(2024, 6, 15)", "Date(2024, 1, 1)"]
    );
}

#[test]
fn test_py_class_staticmethod() {
    let src = r#"
class MathUtils:
    @staticmethod
    def gcd(a, b):
        while b:
            a, b = b, a % b
        return a

    @staticmethod
    def lcm(a, b):
        return a * b // MathUtils.gcd(a, b)

print(MathUtils.gcd(12, 8))
print(MathUtils.lcm(4, 6))
obj = MathUtils()
print(obj.gcd(15, 5))  # also callable via instance
"#;
    assert_eq!(run_python(src), vec!["4", "12", "5"]);
}

#[test]
fn test_py_class_dunder_call() {
    let src = r#"
class Multiplier:
    def __init__(self, factor):
        self.factor = factor

    def __call__(self, x):
        return self.factor * x

triple = Multiplier(3)
print(triple(5))
print(triple(10))
print(callable(triple))
"#;
    assert_eq!(run_python(src), vec!["15", "30", "True"]);
}

#[test]
fn test_py_class_dunder_hash_eq() {
    let src = r#"
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __eq__(self, other):
        return self.x == other.x and self.y == other.y

    def __hash__(self):
        return hash((self.x, self.y))

p1 = Point(1, 2)
p2 = Point(1, 2)
p3 = Point(3, 4)
print(p1 == p2)
print(hash(p1) == hash(p2))
s = {p1, p2, p3}
print(len(s))  # p1 and p2 are equal, so only 2 in set
"#;
    assert_eq!(run_python(src), vec!["True", "True", "2"]);
}

#[test]
fn test_py_class_getattr_setattr_hasattr() {
    let src = r#"
class FlexObj:
    def __getattr__(self, name):
        return f"default_{name}"

    def __setattr__(self, name, value):
        object.__setattr__(self, name, value.upper() if isinstance(value, str) else value)

f = FlexObj()
f.name = "alice"
print(f.name)
print(f.undefined)  # triggers __getattr__
print(hasattr(f, 'name'))
"#;
    assert_eq!(run_python(src), vec!["ALICE", "default_undefined", "True"]);
}

#[test]
fn test_py_class_dunder_enter_exit_protocol() {
    let src = r#"
class ManagedResource:
    def __init__(self, name):
        self.name = name

    def __enter__(self):
        print(f"Acquiring {self.name}")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        print(f"Releasing {self.name}")
        return False  # don't suppress exceptions

with ManagedResource("DB") as r:
    print(f"Using {r.name}")
"#;
    assert_eq!(
        run_python(src),
        vec!["Acquiring DB", "Using DB", "Releasing DB"]
    );
}
