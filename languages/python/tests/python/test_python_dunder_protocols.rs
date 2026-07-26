// Python dunder protocols — __len__, __contains__, __iter__, __getitem__, __call__, __enter__/__exit__
use super::helpers::run_python;

#[test]
fn test_dunder_len() {
    let script = r#"
class Bag:
    def __init__(self, items):
        self.items = items
    def __len__(self):
        return len(self.items)

b = Bag([1, 2, 3])
print(len(b))
"#;
    assert_eq!(run_python(script), vec!["3"]);
}

#[test]
fn test_dunder_contains() {
    let script = r#"
class WordSet:
    def __init__(self, words):
        self.words = set(words)
    def __contains__(self, item):
        return item.lower() in self.words

ws = WordSet(["hello", "world"])
print("hello" in ws)
print("WORLD" in ws)
print("python" in ws)
"#;
    assert_eq!(run_python(script), vec!["True", "True", "False"]);
}

#[test]
fn test_dunder_getitem_setitem() {
    let script = r#"
class Grid:
    def __init__(self):
        self._data = {}
    def __setitem__(self, key, value):
        self._data[key] = value
    def __getitem__(self, key):
        return self._data[key]

g = Grid()
g[(0, 0)] = "A"
g[(1, 1)] = "B"
print(g[(0, 0)])
print(g[(1, 1)])
"#;
    assert_eq!(run_python(script), vec!["A", "B"]);
}

#[test]
fn test_dunder_iter() {
    let script = r#"
class CountUp:
    def __init__(self, limit):
        self.limit = limit
    def __iter__(self):
        self.current = 0
        return self
    def __next__(self):
        if self.current >= self.limit:
            raise StopIteration
        v = self.current
        self.current += 1
        return v

print(list(CountUp(4)))
"#;
    assert_eq!(run_python(script), vec!["[0, 1, 2, 3]"]);
}

#[test]
fn test_dunder_call() {
    let script = r#"
class Multiplier:
    def __init__(self, factor):
        self.factor = factor
    def __call__(self, x):
        return x * self.factor

double = Multiplier(2)
triple = Multiplier(3)
print(double(5))
print(triple(5))
"#;
    assert_eq!(run_python(script), vec!["10", "15"]);
}

#[test]
fn test_dunder_str_repr() {
    let script = r#"
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __str__(self):
        return f"Point({self.x}, {self.y})"
    def __repr__(self):
        return f"Point(x={self.x!r}, y={self.y!r})"

p = Point(1, 2)
print(str(p))
print(repr(p))
"#;
    assert_eq!(run_python(script), vec!["Point(1, 2)", "Point(x=1, y=2)"]);
}

#[test]
fn test_dunder_eq_hash() {
    let script = r#"
class Coord:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __eq__(self, other):
        return self.x == other.x and self.y == other.y
    def __hash__(self):
        return hash((self.x, self.y))

c1 = Coord(1, 2)
c2 = Coord(1, 2)
c3 = Coord(3, 4)
print(c1 == c2)
print(c1 == c3)
s = {c1, c2, c3}
print(len(s))
"#;
    assert_eq!(run_python(script), vec!["True", "False", "2"]);
}

#[test]
fn test_dunder_context_manager() {
    let script = r#"
class Resource:
    def __init__(self, name):
        self.name = name
    def __enter__(self):
        print(f"open {self.name}")
        return self
    def __exit__(self, exc_type, exc_val, exc_tb):
        print(f"close {self.name}")
        return False

with Resource("db") as r:
    print(f"using {r.name}")
"#;
    assert_eq!(run_python(script), vec!["open db", "using db", "close db"]);
}

#[test]
fn test_dunder_bool() {
    let script = r#"
class NonEmpty:
    def __init__(self, data):
        self.data = data
    def __bool__(self):
        return len(self.data) > 0

print(bool(NonEmpty([1, 2])))
print(bool(NonEmpty([])))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}
