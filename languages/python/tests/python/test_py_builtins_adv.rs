use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: builtins — zip, enumerate, sorted, map, filter, any, all, min, max, sum, vars, dir, getattr, isinstance
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_builtins_zip_and_unzip() {
    let src = r#"
keys = ["a", "b", "c"]
values = [1, 2, 3]
pairs = list(zip(keys, values))
print(pairs)
print(dict(pairs))

# Unzip
k, v = zip(*pairs)
print(list(k))
print(list(v))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[('a', 1), ('b', 2), ('c', 3)]",
            "{'a': 1, 'b': 2, 'c': 3}",
            "['a', 'b', 'c']",
            "[1, 2, 3]"
        ]
    );
}

#[test]
fn test_py_builtins_enumerate_with_start() {
    let src = r#"
items = ["apple", "banana", "cherry"]
for idx, item in enumerate(items, start=1):
    print(f"{idx}: {item}")
"#;
    assert_eq!(run_python(src), vec!["1: apple", "2: banana", "3: cherry"]);
}

#[test]
fn test_py_builtins_any_all_short_circuit() {
    let src = r#"
nums = [1, 3, 5, 7]
print(all(n % 2 != 0 for n in nums))
print(any(n > 5 for n in nums))
print(all(n > 5 for n in nums))
print(any(n % 2 == 0 for n in nums))
print(any([]))
print(all([]))  # vacuously true
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "False", "False", "False", "True"]
    );
}

#[test]
fn test_py_builtins_min_max_with_key() {
    let src = r#"
words = ["apple", "fig", "banana", "kiwi"]
print(min(words))
print(max(words))
print(min(words, key=len))
print(max(words, key=len))
print(min(5, 3, 8, 1))
"#;
    assert_eq!(run_python(src), vec!["apple", "kiwi", "fig", "banana", "1"]);
}

#[test]
fn test_py_builtins_sum_and_arithmetic() {
    let src = r#"
print(sum([1, 2, 3, 4, 5]))
print(sum([[1, 2], [3, 4]], []))  # sum for list concatenation (start=[])
print(sum(range(100)))
print(sum([0.1] * 10))  # floating point
"#;
    assert_eq!(
        run_python(src),
        vec!["15", "[1, 2, 3, 4]", "4950", "0.9999999999999999"]
    );
}

#[test]
fn test_py_builtins_sorted_stable() {
    let src = r#"
data = [(1, "b"), (2, "a"), (1, "a"), (2, "b")]
# Stable sort: equal keys preserve original order
by_first = sorted(data, key=lambda x: x[0])
print(by_first)

# Sort with multiple criteria
by_both = sorted(data, key=lambda x: (x[0], x[1]))
print(by_both)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[(1, 'b'), (1, 'a'), (2, 'a'), (2, 'b')]",
            "[(1, 'a'), (1, 'b'), (2, 'a'), (2, 'b')]"
        ]
    );
}

#[test]
fn test_py_builtins_map_filter_chain() {
    let src = r#"
data = range(10)
result = list(map(lambda x: x**2, filter(lambda x: x % 2 == 0, data)))
print(result)

# Equivalent with comprehension:
result2 = [x**2 for x in data if x % 2 == 0]
print(result2)
"#;
    assert_eq!(
        run_python(src),
        vec!["[0, 4, 16, 36, 64]", "[0, 4, 16, 36, 64]"]
    );
}

#[test]
fn test_py_builtins_vars_dir_getattr() {
    let src = r#"
class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

p = Person("Alice", 30)
v = vars(p)
print(v)

d = [x for x in dir(p) if not x.startswith("_")]
print("name" in d)
print("age" in d)

print(getattr(p, "name"))
print(getattr(p, "missing", "default"))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "{'name': 'Alice', 'age': 30}",
            "True",
            "True",
            "Alice",
            "default"
        ]
    );
}

#[test]
fn test_py_builtins_isinstance_issubclass() {
    let src = r#"
class Animal: pass
class Dog(Animal): pass
class Cat(Animal): pass

d = Dog()
print(isinstance(d, Dog))
print(isinstance(d, Animal))
print(isinstance(d, (Cat, Dog)))  # checks multiple types
print(isinstance(d, Cat))
print(issubclass(Dog, Animal))
print(issubclass(Dog, (Cat, Animal)))
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "True", "False", "True", "True"]
    );
}

#[test]
fn test_py_builtins_input_int_float_conversion() {
    let src = r#"
print(int("42"))
print(int("0xFF", 16))
print(int("0b1010", 2))
print(int("777", 8))
print(float("3.14"))
print(float("inf"))
print(float("nan") != float("nan"))
"#;
    assert_eq!(
        run_python(src),
        vec!["42", "255", "10", "511", "3.14", "inf", "True"]
    );
}

#[test]
fn test_py_builtins_type_and_isinstance() {
    let src = r#"
print(type(42).__name__)
print(type("hello").__name__)
print(type([1, 2]).__name__)
print(type({"a": 1}).__name__)
print(type(None).__name__)
print(type(lambda: None).__name__)
"#;
    assert_eq!(
        run_python(src),
        vec!["int", "str", "list", "dict", "NoneType", "function"]
    );
}

#[test]
fn test_py_builtins_hash_and_id() {
    let src = r#"
a = "hello"
b = "hello"
print(hash(a) == hash(b))  # same content, same hash

x = [1, 2, 3]
y = x
print(id(x) == id(y))
print(id(x) != id(list(x)))

print(hash(frozenset([1, 2])) == hash(frozenset([2, 1])))  # order doesn't matter
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_builtins_abs_round_divmod_pow() {
    let src = r#"
print(abs(-42))
print(abs(-3.14))
print(round(3.14159, 2))
print(round(2.5))
print(divmod(17, 5))
print(pow(2, 10))
print(pow(3, 4, 7))  # modular exponentiation
"#;
    assert_eq!(
        run_python(src),
        vec!["42", "3.14", "3.14", "2", "(3, 2)", "1024", "4"]
    );
}

#[test]
fn test_py_builtins_chr_ord_hex_bin_oct() {
    let src = r##"
print(ord("A"))
print(chr(65))
print(hex(255))
print(bin(42))
print(oct(8))
print(format(255, "#010b"))
"##;
    assert_eq!(
        run_python(src),
        vec!["65", "A", "0xff", "0b101010", "0o10", "0b11111111"]
    );
}

#[test]
fn test_py_builtins_iter_and_next() {
    let src = r#"
it = iter([1, 2, 3, 4])
print(next(it))
print(next(it))
print(list(it))  # consume rest

# iter with sentinel
import io
buf = io.StringIO("a\nb\nc\n")
lines = list(iter(buf.readline, ""))
print([l.strip() for l in lines])
"#;
    assert_eq!(run_python(src), vec!["1", "2", "[3, 4]", "['a', 'b', 'c']"]);
}
