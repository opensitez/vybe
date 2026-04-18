use vybec::parser_python::parse;
use vybec::compiler_python::Compiler;

fn compile_ok(src: &str) {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&module);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// Python builtins

#[test] fn builtin_range_1() { compile_ok("x = range(10)\n"); }
#[test] fn builtin_range_2() { compile_ok("x = range(1, 10)\n"); }
#[test] fn builtin_range_3() { compile_ok("x = range(0, 20, 2)\n"); }
#[test] fn builtin_enumerate() { compile_ok("for i, v in enumerate([1,2,3]):\n    print(i, v)\n"); }
#[test] fn builtin_zip() { compile_ok("for a, b in zip([1,2], [3,4]):\n    print(a, b)\n"); }
#[test] fn builtin_sorted() { compile_ok("x = sorted([3, 1, 2])\n"); }
#[test] fn builtin_reversed() { compile_ok("x = reversed([1, 2, 3])\n"); }
#[test] fn builtin_sum() { compile_ok("x = sum([1, 2, 3, 4, 5])\n"); }
#[test] fn builtin_min_list() { compile_ok("x = min([3, 1, 2])\n"); }
#[test] fn builtin_max_list() { compile_ok("x = max([3, 1, 2])\n"); }
#[test] fn builtin_min_args() { compile_ok("x = min(1, 2, 3)\n"); }
#[test] fn builtin_max_args() { compile_ok("x = max(1, 2, 3)\n"); }
#[test] fn builtin_any() { compile_ok("x = any([False, True, False])\n"); }
#[test] fn builtin_all() { compile_ok("x = all([True, True, True])\n"); }
#[test] fn builtin_type() { compile_ok("x = type(42)\n"); }
#[test] fn builtin_isinstance() { compile_ok("x = isinstance(42, int)\n"); }
#[test] fn builtin_bool() { compile_ok("x = bool(0)\ny = bool(1)\n"); }
#[test] fn builtin_round() { compile_ok("x = round(3.7)\n"); }
#[test] fn builtin_chr() { compile_ok("x = chr(65)\n"); }
#[test] fn builtin_ord() { compile_ok("x = ord('A')\n"); }
#[test] fn builtin_list_convert() { compile_ok("x = list('hello')\n"); }
#[test] fn builtin_dict_convert() { compile_ok("x = dict()\n"); }
#[test] fn builtin_set_convert() { compile_ok("x = set([1,1,2,3])\n"); }
#[test] fn builtin_tuple_convert() { compile_ok("x = tuple([1,2,3])\n"); }
#[test] fn builtin_hasattr() { compile_ok("x = hasattr(obj, 'name')\n"); }
#[test] fn builtin_repr() { compile_ok("x = repr(42)\n"); }

// String methods

#[test] fn str_find() { compile_ok("x = 'hello'.find('lo')\n"); }
#[test] fn str_index() { compile_ok("x = 'hello'.index('lo')\n"); }
#[test] fn str_count() { compile_ok("x = 'abcabc'.count('abc')\n"); }
#[test] fn str_lstrip() { compile_ok("x = '  hi  '.lstrip()\n"); }
#[test] fn str_rstrip() { compile_ok("x = '  hi  '.rstrip()\n"); }
#[test] fn str_isdigit() { compile_ok("x = '123'.isdigit()\n"); }
#[test] fn str_format() { compile_ok("x = 'hello {}'.format('world')\n"); }
#[test] fn str_encode() { compile_ok("x = 'hello'.encode()\n"); }
#[test] fn str_center() { compile_ok("x = 'hi'.center(10)\n"); }
#[test] fn str_zfill() { compile_ok("x = '42'.zfill(5)\n"); }
#[test] fn str_ljust() { compile_ok("x = 'hi'.ljust(10)\n"); }

// List methods

#[test] fn list_append() { compile_ok("x = []\nx.append(1)\n"); }
#[test] fn list_pop() { compile_ok("x = [1,2,3]\nx.pop()\n"); }
#[test] fn list_extend() { compile_ok("x = [1]\nx.extend([2,3])\n"); }
#[test] fn list_reverse() { compile_ok("x = [1,2,3]\nx.reverse()\n"); }
#[test] fn list_sort() { compile_ok("x = [3,1,2]\nx.sort()\n"); }
#[test] fn list_copy() { compile_ok("x = [1,2,3]\ny = x.copy()\n"); }
#[test] fn list_clear() { compile_ok("x = [1,2,3]\nx.clear()\n"); }

// Dict methods

#[test] fn dict_keys() { compile_ok("x = {'a': 1}.keys()\n"); }
#[test] fn dict_values() { compile_ok("x = {'a': 1}.values()\n"); }
#[test] fn dict_items() { compile_ok("for k, v in d.items():\n    print(k, v)\n"); }
#[test] fn dict_get() { compile_ok("x = d.get('key')\n"); }

// Slicing

#[test] fn slice_basic() { compile_ok("x = [1,2,3,4,5]\ny = x[1:3]\n"); }
#[test] fn slice_from_start() { compile_ok("x = [1,2,3][:2]\n"); }
#[test] fn slice_to_end() { compile_ok("x = [1,2,3][1:]\n"); }
#[test] fn slice_step() { compile_ok("x = [1,2,3,4,5][::2]\n"); }

// In operator

#[test] fn in_list() { compile_ok("x = 1 in [1, 2, 3]\n"); }
#[test] fn not_in_list() { compile_ok("x = 4 not in [1, 2, 3]\n"); }

// With statement

#[test] fn with_basic() { compile_ok("with open('file.txt') as f:\n    data = f.read()\n"); }

// Assert

#[test] fn assert_simple() { compile_ok("assert True\n"); }
#[test] fn assert_msg() { compile_ok("assert x > 0, 'must be positive'\n"); }

// Delete

#[test] fn del_var() { compile_ok("x = 1\ndel x\n"); }

// Augmented assign to attr/subscript

#[test] fn aug_assign_attr() { compile_ok("class C:\n    def inc(self):\n        self.count += 1\n"); }

// Type annotation

#[test] fn type_annotation() { compile_ok("x: int = 5\ny: str = 'hello'\n"); }

// Real-world patterns

#[test]
fn word_frequency() {
    compile_ok(r#"
text = "the cat sat on the mat the cat"
words = text.split()
freq = {}
for w in words:
    if w in freq:
        freq[w] += 1
    else:
        freq[w] = 1
for k, v in freq.items():
    print(f"{k}: {v}")
"#);
}

#[test]
fn enumerate_pattern() {
    compile_ok(r#"
items = ["apple", "banana", "cherry"]
for i, item in enumerate(items):
    print(f"{i}: {item}")
"#);
}

#[test]
fn sorted_with_comp() {
    compile_ok(r#"
data = [5, 2, 8, 1, 9, 3]
ascending = sorted(data)
print(ascending)
total = sum(data)
avg = total / len(data)
print(f"sum={total}, avg={avg}")
"#);
}

#[test]
fn class_with_augmented_attr() {
    compile_ok(r#"
class Counter:
    def __init__(self):
        self.count = 0
    def increment(self):
        self.count += 1
    def get(self):
        return self.count
"#);
}

#[test]
fn try_except_assert() {
    compile_ok(r#"
def safe_div(a, b):
    assert b != 0, "cannot divide by zero"
    return a / b

try:
    result = safe_div(10, 0)
except:
    print("caught error")
"#);
}

#[test]
fn with_and_string_methods() {
    compile_ok(r#"
text = "  Hello, World!  "
cleaned = text.strip().lower()
words = cleaned.split()
print(len(words))
print("hello" in cleaned)
"#);
}

#[test]
fn nested_comprehension_real() {
    compile_ok(r#"
matrix = [[1,2,3],[4,5,6],[7,8,9]]
flat = [x for row in matrix for x in row]
evens = [x for x in flat if x % 2 == 0]
print(sorted(evens))
print(sum(evens))
"#);
}
