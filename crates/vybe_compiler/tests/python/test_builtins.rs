use super::helpers::{compile_ok, run_python_one};

// Python builtins

#[test]
fn builtin_range_1() {
    compile_ok("x = range(10)\n");
}
#[test]
fn builtin_range_2() {
    compile_ok("x = range(1, 10)\n");
}
#[test]
fn builtin_range_3() {
    compile_ok("x = range(0, 20, 2)\n");
}
#[test]
fn builtin_enumerate() {
    compile_ok("for i, v in enumerate([1,2,3]):\n    print(i, v)\n");
}
#[test]
fn builtin_zip() {
    compile_ok("for a, b in zip([1,2], [3,4]):\n    print(a, b)\n");
}
#[test]
fn builtin_sorted() {
    compile_ok("x = sorted([3, 1, 2])\n");
}
#[test]
fn builtin_reversed() {
    compile_ok("x = reversed([1, 2, 3])\n");
}
#[test]
fn builtin_sum() {
    compile_ok("x = sum([1, 2, 3, 4, 5])\n");
}
#[test]
fn builtin_min_list() {
    compile_ok("x = min([3, 1, 2])\n");
}
#[test]
fn builtin_max_list() {
    compile_ok("x = max([3, 1, 2])\n");
}
#[test]
fn builtin_min_args() {
    compile_ok("x = min(1, 2, 3)\n");
}
#[test]
fn builtin_max_args() {
    compile_ok("x = max(1, 2, 3)\n");
}
#[test]
fn builtin_any() {
    compile_ok("x = any([False, True, False])\n");
}
#[test]
fn builtin_all() {
    compile_ok("x = all([True, True, True])\n");
}

// ── Runtime tests for Python any/all (stdlib:pyany / stdlib:pyall polyfills) ──
#[test]
fn builtin_any_runtime_true() {
    assert_eq!(run_python_one("print(any([False, True, False]))\n"), "true");
}
#[test]
fn builtin_any_runtime_false() {
    assert_eq!(
        run_python_one("print(any([False, False, False]))\n"),
        "false"
    );
}
#[test]
fn builtin_any_empty() {
    assert_eq!(run_python_one("print(any([]))\n"), "false");
}
#[test]
fn builtin_all_runtime_true() {
    assert_eq!(run_python_one("print(all([True, 1, 'x']))\n"), "true");
}
#[test]
fn builtin_all_runtime_false() {
    assert_eq!(run_python_one("print(all([True, False, True]))\n"), "false");
}
#[test]
fn builtin_all_empty() {
    assert_eq!(run_python_one("print(all([]))\n"), "true");
}

// ── Runtime tests for new bare-form polyfills ──
#[test]
fn builtin_pymap_runtime() {
    assert_eq!(
        run_python_one("r = list(map(lambda x: x * 2, [1, 2, 3]))\nprint(r[1])\n"),
        "4"
    );
}
#[test]
fn builtin_pyfilter_runtime() {
    assert_eq!(
        run_python_one("r = list(filter(lambda x: x > 1, [1, 2, 3]))\nprint(len(r))\n"),
        "2"
    );
}
#[test]
fn builtin_next_with_default() {
    // Empty iter → default returned
    assert_eq!(run_python_one("print(next([], 'x'))\n"), "x");
}
#[test]
fn builtin_next_consumes_first() {
    assert_eq!(
        run_python_one("a = [10, 20, 30]\nprint(next(a, 0))\n"),
        "10"
    );
}
#[test]
fn builtin_random_choice_in_range() {
    // Single-element list — choice is deterministic
    assert_eq!(
        run_python_one("import random\nprint(random.choice([42]))\n"),
        "42"
    );
}
#[test]
fn builtin_random_shuffle_preserves_count() {
    // Shuffle returns same set; check length unchanged
    assert_eq!(
        run_python_one("import random\na = [1,2,3,4,5]\nrandom.shuffle(a)\nprint(len(a))\n"),
        "5"
    );
}
#[test]
fn builtin_random_sample_size() {
    assert_eq!(
        run_python_one("import random\nr = random.sample([1,2,3,4,5], 3)\nprint(len(r))\n"),
        "3"
    );
}
#[test]
fn builtin_type() {
    compile_ok("x = type(42)\n");
}
#[test]
fn builtin_isinstance() {
    compile_ok("x = isinstance(42, int)\n");
}
#[test]
fn builtin_bool() {
    compile_ok("x = bool(0)\ny = bool(1)\n");
}
#[test]
fn builtin_round() {
    compile_ok("x = round(3.7)\n");
}
#[test]
fn builtin_chr() {
    compile_ok("x = chr(65)\n");
}
#[test]
fn builtin_ord() {
    compile_ok("x = ord('A')\n");
}
#[test]
fn builtin_list_convert() {
    compile_ok("x = list('hello')\n");
}
#[test]
fn builtin_dict_convert() {
    compile_ok("x = dict()\n");
}
#[test]
fn builtin_set_convert() {
    compile_ok("x = set([1,1,2,3])\n");
}
#[test]
fn builtin_tuple_convert() {
    compile_ok("x = tuple([1,2,3])\n");
}
#[test]
fn builtin_hasattr() {
    compile_ok("x = hasattr(obj, 'name')\n");
}
#[test]
fn builtin_repr() {
    compile_ok("x = repr(42)\n");
}
#[test]
fn builtin_abs() {
    compile_ok("x = abs(-5)\n");
}
#[test]
fn builtin_len() {
    compile_ok("x = len([1,2,3])\n");
}
#[test]
fn builtin_int() {
    compile_ok("x = int('5')\n");
}
#[test]
fn builtin_float() {
    compile_ok("x = float('3.14')\n");
}
#[test]
fn builtin_str() {
    compile_ok("x = str(42)\n");
}
#[test]
fn builtin_hex() {
    compile_ok("s = hex(255)\n");
}
#[test]
fn builtin_oct() {
    compile_ok("s = oct(8)\n");
}
#[test]
fn builtin_bin() {
    compile_ok("s = bin(42)\n");
}
#[test]
fn builtin_id() {
    compile_ok("x = [1, 2]\nn = id(x)\n");
}
#[test]
fn builtin_hash() {
    compile_ok("n = hash('hello')\n");
}
#[test]
fn builtin_callable() {
    compile_ok("def f(): pass\nb = callable(f)\n");
}
#[test]
fn builtin_frozenset() {
    compile_ok("s = frozenset([1, 2, 3])\n");
}
#[test]
fn builtin_vars() {
    compile_ok("class C:\n    def __init__(self):\n        self.x = 1\nc = C()\nd = vars(c)\n");
}
#[test]
fn builtin_dir() {
    compile_ok("x = {'a': 1, 'b': 2}\nkeys = dir(x)\n");
}
#[test]
fn builtin_pow_two() {
    compile_ok("x = pow(2, 10)\n");
}
#[test]
fn builtin_pow_three() {
    compile_ok("x = pow(2, 10, 100)\n");
}
#[test]
fn builtin_divmod() {
    compile_ok("q, r = divmod(17, 5)\n");
}
#[test]
fn builtin_format() {
    compile_ok("s = format(3.14, '.2f')\n");
}
#[test]
fn builtin_iter() {
    compile_ok("it = iter([1, 2, 3])\n");
}
#[test]
fn builtin_next() {
    compile_ok("items = [1, 2, 3]\nx = next(items)\n");
}

// String methods

#[test]
fn str_upper() {
    compile_ok("x = 'hello'.upper()\n");
}
#[test]
fn str_lower() {
    compile_ok("x = 'HELLO'.lower()\n");
}
#[test]
fn str_strip() {
    compile_ok("x = '  hi  '.strip()\n");
}
#[test]
fn str_split() {
    compile_ok("x = 'a b c'.split()\n");
}
#[test]
fn str_join() {
    compile_ok("x = ' '.join(['a', 'b', 'c'])\n");
}
#[test]
fn str_replace() {
    compile_ok("x = 'hello'.replace('l', 'r')\n");
}
#[test]
fn str_find() {
    compile_ok("x = 'hello'.find('lo')\n");
}
#[test]
fn str_index() {
    compile_ok("x = 'hello'.index('lo')\n");
}
#[test]
fn str_count() {
    compile_ok("x = 'abcabc'.count('abc')\n");
}
#[test]
fn str_startswith() {
    compile_ok("x = 'hello'.startswith('he')\n");
}
#[test]
fn str_endswith() {
    compile_ok("x = 'hello'.endswith('lo')\n");
}
#[test]
fn str_lstrip() {
    compile_ok("x = '  hi  '.lstrip()\n");
}
#[test]
fn str_rstrip() {
    compile_ok("x = '  hi  '.rstrip()\n");
}
#[test]
fn str_isdigit() {
    compile_ok("x = '123'.isdigit()\n");
}
#[test]
fn str_format() {
    compile_ok("x = 'hello {}'.format('world')\n");
}
#[test]
fn str_encode() {
    compile_ok("x = 'hello'.encode()\n");
}
#[test]
fn str_center() {
    compile_ok("x = 'hi'.center(10)\n");
}
#[test]
fn str_zfill() {
    compile_ok("x = '42'.zfill(5)\n");
}
#[test]
fn str_ljust() {
    compile_ok("x = 'hi'.ljust(10)\n");
}
#[test]
fn str_capitalize() {
    compile_ok("x = 'hello'.capitalize()\n");
}
#[test]
fn str_title() {
    compile_ok("x = 'hello world'.title()\n");
}
#[test]
fn str_swapcase() {
    compile_ok("x = 'Hello'.swapcase()\n");
}
#[test]
fn str_removeprefix() {
    compile_ok("s = 'HelloWorld'\nr = s.removeprefix('Hello')\n");
}
#[test]
fn str_removesuffix() {
    compile_ok("s = 'HelloWorld'\nr = s.removesuffix('World')\n");
}

// List methods

#[test]
fn list_append() {
    compile_ok("x = []\nx.append(1)\n");
}
#[test]
fn list_pop() {
    compile_ok("x = [1,2,3]\nx.pop()\n");
}
#[test]
fn list_extend() {
    compile_ok("x = [1]\nx.extend([2,3])\n");
}
#[test]
fn list_reverse() {
    compile_ok("x = [1,2,3]\nx.reverse()\n");
}
#[test]
fn list_sort() {
    compile_ok("x = [3,1,2]\nx.sort()\n");
}
#[test]
fn list_copy() {
    compile_ok("x = [1,2,3]\ny = x.copy()\n");
}
#[test]
fn list_clear() {
    compile_ok("x = [1,2,3]\nx.clear()\n");
}
#[test]
fn list_insert() {
    compile_ok("x = [1,3]\nx.insert(1, 2)\n");
}
#[test]
fn list_remove() {
    compile_ok("x = [1,2,3]\nx.remove(2)\n");
}

// Dict methods

#[test]
fn dict_keys() {
    compile_ok("x = {'a': 1}.keys()\n");
}
#[test]
fn dict_values() {
    compile_ok("x = {'a': 1}.values()\n");
}
#[test]
fn dict_items() {
    compile_ok("for k, v in d.items():\n    print(k, v)\n");
}
#[test]
fn dict_get() {
    compile_ok("x = d.get('key')\n");
}
#[test]
fn dict_update() {
    compile_ok("d = {'a': 1}\nd.update({'b': 2})\n");
}
#[test]
fn dict_setdefault() {
    compile_ok("d = {}\nv = d.setdefault('key', 42)\n");
}

// Set methods

#[test]
fn set_add() {
    compile_ok("s = {1, 2, 3}\ns.add(4)\n");
}
#[test]
fn set_discard() {
    compile_ok("s = {1, 2, 3}\ns.discard(2)\n");
}
#[test]
fn set_union() {
    compile_ok("a = {1, 2}\nb = {3, 4}\nc = a.union(b)\n");
}
#[test]
fn set_intersection() {
    compile_ok("a = {1, 2, 3}\nb = {2, 3, 4}\nc = a.intersection(b)\n");
}
#[test]
fn set_difference() {
    compile_ok("a = {1, 2, 3}\nb = {2, 3}\nc = a.difference(b)\n");
}

// Slicing

#[test]
fn slice_basic() {
    compile_ok("x = [1,2,3,4,5]\ny = x[1:3]\n");
}
#[test]
fn slice_from_start() {
    compile_ok("x = [1,2,3][:2]\n");
}
#[test]
fn slice_to_end() {
    compile_ok("x = [1,2,3][1:]\n");
}
#[test]
fn slice_step() {
    compile_ok("x = [1,2,3,4,5][::2]\n");
}
#[test]
fn slice_step_reverse() {
    compile_ok("x = [1,2,3,4,5]\ny = x[::-1]\n");
}
#[test]
fn slice_step_with_bounds() {
    compile_ok("x = [1,2,3,4,5]\ny = x[1:4:2]\n");
}
#[test]
fn slice_step_string() {
    compile_ok("s = 'hello'\nr = s[::-1]\n");
}

// In operator

#[test]
fn in_list() {
    compile_ok("x = 1 in [1, 2, 3]\n");
}
#[test]
fn not_in_list() {
    compile_ok("x = 4 not in [1, 2, 3]\n");
}

// With statement

#[test]
fn with_basic() {
    compile_ok("with open('file.txt') as f:\n    data = f.read()\n");
}

// Assert

#[test]
fn assert_simple() {
    compile_ok("assert True\n");
}
#[test]
fn assert_msg() {
    compile_ok("assert x > 0, 'must be positive'\n");
}

// Delete

#[test]
fn del_var() {
    compile_ok("x = 1\ndel x\n");
}
#[test]
fn del_dict_key() {
    compile_ok("d = {'a': 1, 'b': 2}\ndel d['a']\n");
}
#[test]
fn del_list_index() {
    compile_ok("lst = [1, 2, 3]\ndel lst[0]\n");
}
#[test]
fn del_attribute() {
    compile_ok("class C:\n    pass\nc = C()\nc.x = 10\ndel c.x\n");
}
#[test]
fn del_multiple() {
    compile_ok("a = 1\nb = 2\ndel a, b\n");
}

// String repetition

#[test]
fn str_repeat_basic() {
    compile_ok("x = '=' * 40\n");
}
#[test]
fn str_repeat_int_first() {
    compile_ok("x = 3 * 'abc'\n");
}

// Sorted with key

#[test]
fn sorted_with_key() {
    compile_ok("words = ['banana', 'apple', 'cherry']\nresult = sorted(words, key=len)\n");
}
#[test]
fn sorted_with_lambda_key() {
    compile_ok(
        "pairs = [(1, 'b'), (3, 'a'), (2, 'c')]\nresult = sorted(pairs, key=lambda x: x[1])\n",
    );
}

// f-string format specs

#[test]
fn fstring_format_float() {
    compile_ok("x = 3.14159\ns = f\"{x:.2f}\"\n");
}
#[test]
fn fstring_format_width() {
    compile_ok("name = \"hello\"\ns = f\"{name:>20}\"\n");
}
#[test]
fn fstring_format_int() {
    compile_ok("n = 42\ns = f\"{n:04d}\"\n");
}

// Extended unpacking

#[test]
fn list_unpack_star() {
    compile_ok("a = [2, 3]\nb = [1, *a, 4]\n");
}
#[test]
fn list_unpack_multiple() {
    compile_ok("x = [1, 2]\ny = [3, 4]\nz = [*x, *y]\n");
}
#[test]
fn tuple_unpack_star() {
    compile_ok("a = (1, 2)\nb = (*a, 3, 4)\n");
}

// Slice assignment

#[test]
fn slice_assign_basic() {
    compile_ok("a = [1, 2, 3, 4, 5]\na[1:3] = [10, 20]\n");
}
#[test]
fn slice_assign_full() {
    compile_ok("a = [1, 2, 3]\na[:] = [4, 5, 6]\n");
}

// Augmented assign to attr/subscript

#[test]
fn aug_assign_attr() {
    compile_ok("class C:\n    def inc(self):\n        self.count += 1\n");
}
