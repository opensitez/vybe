use super::helpers::run_python;

#[test]
fn test_numbers_abc_imports() {
    let script = r#"
import numbers

print(hasattr(numbers, 'Number'))
print(hasattr(numbers, 'Complex'))
print(hasattr(numbers, 'Real'))
print(hasattr(numbers, 'Rational'))
print(hasattr(numbers, 'Integral'))
"#;
    assert_eq!(
        run_python(script),
        vec!["True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_numbers_int_hierarchy() {
    let script = r#"
import numbers

x = 42
print(isinstance(x, numbers.Number))
print(isinstance(x, numbers.Complex))
print(isinstance(x, numbers.Real))
print(isinstance(x, numbers.Rational))
print(isinstance(x, numbers.Integral))
"#;
    assert_eq!(
        run_python(script),
        vec!["True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_numbers_float_hierarchy() {
    let script = r#"
import numbers

x = 3.14
print(isinstance(x, numbers.Number))
print(isinstance(x, numbers.Complex))
print(isinstance(x, numbers.Real))
print(isinstance(x, numbers.Rational))
print(isinstance(x, numbers.Integral))
"#;
    assert_eq!(
        run_python(script),
        vec!["True", "True", "True", "False", "False"]
    );
}

#[test]
fn test_numbers_complex_hierarchy() {
    let script = r#"
import numbers

x = 1 + 2j
print(isinstance(x, numbers.Number))
print(isinstance(x, numbers.Complex))
print(isinstance(x, numbers.Real))
print(isinstance(x, numbers.Rational))
print(isinstance(x, numbers.Integral))
"#;
    assert_eq!(
        run_python(script),
        vec!["True", "True", "False", "False", "False"]
    );
}

#[test]
fn test_numbers_fraction_hierarchy() {
    let script = r#"
import numbers
from fractions import Fraction

x = Fraction(3, 4)
print(isinstance(x, numbers.Number))
print(isinstance(x, numbers.Complex))
print(isinstance(x, numbers.Real))
print(isinstance(x, numbers.Rational))
print(isinstance(x, numbers.Integral))
"#;
    assert_eq!(
        run_python(script),
        vec!["True", "True", "True", "True", "False"]
    );
}

#[test]
fn test_numbers_issubclass_relationships() {
    let script = r#"
import numbers

print(issubclass(numbers.Integral, numbers.Rational))
print(issubclass(numbers.Rational, numbers.Real))
print(issubclass(numbers.Real, numbers.Complex))
print(issubclass(numbers.Complex, numbers.Number))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_numbers_bool_is_integral() {
    let script = r#"
import numbers

print(isinstance(True, numbers.Integral))
print(isinstance(False, numbers.Integral))
print(issubclass(bool, numbers.Integral))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}

#[test]
fn test_numbers_register_custom_type() {
    let script = r#"
import numbers

class MyNumber:
    pass

numbers.Number.register(MyNumber)

obj = MyNumber()
print(isinstance(obj, numbers.Number))
print(issubclass(MyNumber, numbers.Number))
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_numbers_register_custom_real() {
    let script = r#"
import numbers

class MyReal:
    pass

numbers.Real.register(MyReal)

obj = MyReal()
print(isinstance(obj, numbers.Real))
print(isinstance(obj, numbers.Complex))
print(isinstance(obj, numbers.Number))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}

#[test]
fn test_numbers_subclass_integral_implementation() {
    let script = r#"
import numbers

class CustomInt(numbers.Integral):
    def __init__(self, val):
        self.val = val
    def __int__(self):
        return int(self.val)
    def __index__(self):
        return int(self.val)
    def __abs__(self):
        return abs(self.val)
    def __add__(self, other):
        return self.val + other
    def __radd__(self, other):
        return other + self.val
    def __sub__(self, other):
        return self.val - other
    def __rsub__(self, other):
        return other - self.val
    def __mul__(self, other):
        return self.val * other
    def __rmul__(self, other):
        return other * self.val
    def __truediv__(self, other):
        return self.val / other
    def __rtruediv__(self, other):
        return other / self.val
    def __floordiv__(self, other):
        return self.val // other
    def __rfloordiv__(self, other):
        return other // self.val
    def __mod__(self, other):
        return self.val % other
    def __rmod__(self, other):
        return other % self.val
    def __pow__(self, other):
        return self.val ** other
    def __rpow__(self, other):
        return other ** self.val
    def __rshift__(self, other):
        return self.val >> other
    def __rrshift__(self, other):
        return other >> self.val
    def __lshift__(self, other):
        return self.val << other
    def __rlshift__(self, other):
        return other << self.val
    def __and__(self, other):
        return self.val & other
    def __rand__(self, other):
        return other & self.val
    def __or__(self, other):
        return self.val | other
    def __ror__(self, other):
        return other | self.val
    def __xor__(self, other):
        return self.val ^ other
    def __rxor__(self, other):
        return other ^ self.val
    def __neg__(self):
        return -self.val
    def __pos__(self):
        return +self.val
    def __invert__(self):
        return ~self.val
    def __eq__(self, other):
        return self.val == other
    def __lt__(self, other):
        return self.val < other
    def __le__(self, other):
        return self.val <= other
    def __trunc__(self):
        return self.val
    def __floor__(self):
        return self.val
    def __ceil__(self):
        return self.val
    def __round__(self, n=None):
        return round(self.val, n)

c = CustomInt(10)
print(isinstance(c, numbers.Integral))
print(int(c) == 10)
print(c + 5 == 15)
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}

#[test]
fn test_numbers_non_numeric_type_checks() {
    let script = r#"
import numbers

print(isinstance("123", numbers.Number))
print(isinstance([1, 2, 3], numbers.Number))
print(isinstance({'a': 1}, numbers.Number))
print(isinstance((1, 2), numbers.Number))
print(isinstance(None, numbers.Number))
"#;
    assert_eq!(
        run_python(script),
        vec!["False", "False", "False", "False", "False"]
    );
}

#[test]
fn test_numbers_real_properties() {
    let script = r#"
import numbers

x = 5.5
print(hasattr(x, 'real'))
print(hasattr(x, 'imag'))
print(x.real == 5.5)
print(x.imag == 0.0)
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_numbers_complex_properties() {
    let script = r#"
import numbers

z = 3 + 4j
print(z.real)
print(z.imag)
print(z.conjugate() == (3 - 4j))
"#;
    assert_eq!(run_python(script), vec!["3.0", "4.0", "True"]);
}

#[test]
fn test_numbers_rational_numerator_denominator() {
    let script = r#"
import numbers
from fractions import Fraction

f = Fraction(5, 8)
print(isinstance(f, numbers.Rational))
print(f.numerator)
print(f.denominator)

i = 7
print(isinstance(i, numbers.Rational))
print(i.numerator)
print(i.denominator)
"#;
    assert_eq!(run_python(script), vec!["True", "5", "8", "True", "7", "1"]);
}

#[test]
fn test_numbers_decimal_not_registered_by_default() {
    let script = r#"
import numbers
from decimal import Decimal

d = Decimal('3.14')
print(isinstance(d, numbers.Number))
print(isinstance(d, numbers.Real))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_numbers_virtual_subclass_check() {
    let script = r#"
import numbers

class MyIntLike:
    def __int__(self):
        return 1

numbers.Integral.register(MyIntLike)
print(issubclass(MyIntLike, numbers.Integral))
print(issubclass(MyIntLike, numbers.Number))
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_numbers_abc_dir_and_all() {
    let script = r#"
import numbers

names = dir(numbers)
for expected in ['Number', 'Complex', 'Real', 'Rational', 'Integral']:
    print(expected in names)
"#;
    assert_eq!(
        run_python(script),
        vec!["True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_numbers_complex_abs_and_conjugate() {
    let script = r#"
import numbers

z = 3 + 4j
print(abs(z))
print((z * z.conjugate()).real)
"#;
    assert_eq!(run_python(script), vec!["5.0", "25.0"]);
}

#[test]
fn test_numbers_integral_bitwise_operations() {
    let script = r#"
import numbers

a = 0b1010
b = 0b1100

print(isinstance(a & b, numbers.Integral))
print(isinstance(a | b, numbers.Integral))
print(isinstance(a ^ b, numbers.Integral))
print(isinstance(~a, numbers.Integral))
print(isinstance(a << 2, numbers.Integral))
print(isinstance(b >> 1, numbers.Integral))
"#;
    assert_eq!(
        run_python(script),
        vec!["True", "True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_numbers_abc_instantiation_raises() {
    let script = r#"
import numbers

try:
    numbers.Number()
    print('no_error')
except TypeError:
    print('TypeError_raised')
"#;
    assert_eq!(run_python(script), vec!["TypeError_raised"]);
}
