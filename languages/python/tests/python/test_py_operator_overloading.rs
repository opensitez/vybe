use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Operator Overloading — dunder methods for arithmetic, in-place, reflected, indexing, container ops
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_dunder_add_radd_iadd() {
    let src = r#"
class Number:
    def __init__(self, val):
        self.val = val

    def __add__(self, other):
        v = other.val if isinstance(other, Number) else other
        return Number(self.val + v)

    def __radd__(self, other):
        return self.__add__(other)

    def __iadd__(self, other):
        v = other.val if isinstance(other, Number) else other
        self.val += v
        return self

    def __repr__(self):
        return f"Num({self.val})"

n1 = Number(10)
n2 = Number(20)
print(n1 + n2)
print(100 + n1)
n1 += 5
print(n1)
"#;
    assert_eq!(run_python(src), vec!["Num(30)", "Num(110)", "Num(15)"]);
}

#[test]
fn test_py_dunder_sub_mul_truediv_floordiv_mod() {
    let src = r#"
class Quantity:
    def __init__(self, val):
        self.val = val

    def __sub__(self, other): return Quantity(self.val - other)
    def __mul__(self, other): return Quantity(self.val * other)
    def __truediv__(self, other): return Quantity(self.val / other)
    def __floordiv__(self, other): return Quantity(self.val // other)
    def __mod__(self, other): return Quantity(self.val % other)
    def __repr__(self): return f"Q({self.val})"

q = Quantity(20)
print(q - 5)
print(q * 3)
print(q / 4)
print(q // 3)
print(q % 3)
"#;
    assert_eq!(
        run_python(src),
        vec!["Q(15)", "Q(60)", "Q(5.0)", "Q(6)", "Q(2)"]
    );
}

#[test]
fn test_py_dunder_getitem_setitem_delitem_slice() {
    let src = r#"
class Grid:
    def __init__(self):
        self._data = {}

    def __getitem__(self, key):
        return self._data.get(key, 0)

    def __setitem__(self, key, val):
        self._data[key] = val

    def __delitem__(self, key):
        del self._data[key]

g = Grid()
g[1, 2] = 100
g[3, 4] = 200
print(g[1, 2])
print(g[0, 0])
del g[1, 2]
print(g[1, 2])
"#;
    assert_eq!(run_python(src), vec!["100", "0", "0"]);
}

#[test]
fn test_py_dunder_contains_membership() {
    let src = r#"
class CustomRange:
    def __init__(self, low, high):
        self.low = low
        self.high = high

    def __contains__(self, val):
        return self.low <= val <= self.high

r = CustomRange(10, 50)
print(25 in r)
print(5 in r)
print(50 in r)
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True"]);
}

#[test]
fn test_py_dunder_call_invocable_object() {
    let src = r#"
class Polynomial:
    def __init__(self, a, b, c):
        self.a = a
        self.b = b
        self.c = c

    def __call__(self, x):
        return self.a * x**2 + self.b * x + self.c

p = Polynomial(2, 3, 1)
print(p(0))
print(p(2))
print(callable(p))
"#;
    assert_eq!(run_python(src), vec!["1", "15", "True"]);
}

#[test]
fn test_py_dunder_bitwise_operators() {
    let src = r#"
class BitSet:
    def __init__(self, mask):
        self.mask = mask

    def __and__(self, other): return BitSet(self.mask & other.mask)
    def __or__(self, other): return BitSet(self.mask | other.mask)
    def __xor__(self, other): return BitSet(self.mask ^ other.mask)
    def __invert__(self): return BitSet(~self.mask)
    def __repr__(self): return f"BitSet({bin(self.mask)})"

b1 = BitSet(0b1100)
b2 = BitSet(0b1010)
print(b1 & b2)
print(b1 | b2)
print(b1 ^ b2)
"#;
    assert_eq!(
        run_python(src),
        vec!["BitSet(0b1000)", "BitSet(0b1110)", "BitSet(0b0110)"]
    );
}

#[test]
fn test_py_dunder_unary_neg_pos_abs() {
    let src = r#"
class Point2D:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __neg__(self): return Point2D(-self.x, -self.y)
    def __pos__(self): return Point2D(self.x, self.y)
    def __abs__(self): return (self.x**2 + self.y**2)**0.5
    def __repr__(self): return f"Point({self.x}, {self.y})"

p = Point2D(3, -4)
print(-p)
print(abs(p))
"#;
    assert_eq!(run_python(src), vec!["Point(-3, 4)", "5.0"]);
}

#[test]
fn test_py_not_implemented_fallthrough() {
    let src = r#"
class A:
    def __add__(self, other):
        return NotImplemented

class B:
    def __radd__(self, other):
        return "B handled add"

a = A()
b = B()
print(a + b)
"#;
    assert_eq!(run_python(src), vec!["B handled add"]);
}

#[test]
fn test_py_dunder_matmul_at_operator() {
    let src = r#"
class Matrix1D:
    def __init__(self, data):
        self.data = data

    def __matmul__(self, other):
        return sum(x * y for x, y in zip(self.data, other.data))

m1 = Matrix1D([1, 2, 3])
m2 = Matrix1D([4, 5, 6])
print(m1 @ m2)
"#;
    assert_eq!(run_python(src), vec!["32"]);
}

#[test]
fn test_py_dunder_enter_exit_context_manager_custom() {
    let src = r#"
class Transaction:
    def __enter__(self):
        print("BEGIN")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            print("ROLLBACK")
            return True  # suppress
        print("COMMIT")
        return False

with Transaction():
    print("WORK")

with Transaction():
    print("WORK FAIL")
    raise ValueError("error")
"#;
    assert_eq!(
        run_python(src),
        vec!["BEGIN", "WORK", "COMMIT", "BEGIN", "WORK FAIL", "ROLLBACK"]
    );
}
