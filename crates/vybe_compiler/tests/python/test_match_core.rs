use crate::helpers::run_python_one;

#[test]
fn match_literal_int() {
    assert_eq!(
        run_python_one("x = 1\nmatch x:\n case 1:\n  print('one')\n"),
        "one"
    );
}

#[test]
fn match_literal_string() {
    assert_eq!(
        run_python_one("x = 'a'\nmatch x:\n case 'a':\n  print('letter')\n"),
        "letter"
    );
}

#[test]
fn match_wildcard_default() {
    assert_eq!(
        run_python_one("x = 9\nmatch x:\n case 1:\n  print('no')\n case _:\n  print('other')\n"),
        "other"
    );
}

#[test]
fn match_or_pattern() {
    assert_eq!(
        run_python_one("x = 2\nmatch x:\n case 1 | 2:\n  print('small')\n"),
        "small"
    );
}

#[test]
fn match_sequence_unpack() {
    assert_eq!(
        run_python_one("x = [1, 2]\nmatch x:\n case [a, b]:\n  print(a + b)\n"),
        "3"
    );
}

#[test]
fn match_sequence_with_guard() {
    assert_eq!(
        run_python_one("x = [3, 4]\nmatch x:\n case [a, b] if a < b:\n  print('ok')\n"),
        "ok"
    );
}

#[test]
fn match_mapping_exact_key() {
    assert_eq!(
        run_python_one("d = {'k': 7}\nmatch d:\n case {'k': v}:\n  print(v)\n"),
        "7"
    );
}

#[test]
fn match_class_pattern_positional() {
    assert_eq!(
        run_python_one(
            "class P:\n def __init__(self, x, y):\n  self.x = x\n  self.y = y\np = P(1, 2)\nmatch p:\n case P(x, y):\n  print(x, y)\n"
        ),
        "1 2"
    );
}

#[test]
fn match_as_binding() {
    assert_eq!(
        run_python_one("x = 5\nmatch x:\n case n as value:\n  print(value)\n"),
        "5"
    );
}

#[test]
fn match_nested_sequence() {
    assert_eq!(
        run_python_one("x = (1, (2, 3))\nmatch x:\n case (a, (b, c)):\n  print(b, c)\n"),
        "2 3"
    );
}

#[test]
fn match_tuple_pattern() {
    assert_eq!(
        run_python_one("t = (1, 2)\nmatch t:\n case (x, y):\n  print(x * y)\n"),
        "2"
    );
}

#[test]
fn match_first_case_wins() {
    assert_eq!(
        run_python_one(
            "x = 1\nmatch x:\n case 1:\n  print('first')\n case 1:\n  print('second')\n"
        ),
        "first"
    );
}

#[test]
fn match_no_case_raises() {
    assert_eq!(
        run_python_one(
            "x = 1\nmatch x:\n case 2:\n  print('no')\ntry:\n match x:\n  case 2:\n   pass\nexcept:\n print('unmatched')\n"
        ),
        "unmatched"
    );
}

#[test]
fn match_bool_true() {
    assert_eq!(
        run_python_one("x = True\nmatch x:\n case True:\n  print('t')\n"),
        "t"
    );
}

#[test]
fn match_none_singleton() {
    assert_eq!(
        run_python_one("x = None\nmatch x:\n case None:\n  print('nil')\n"),
        "nil"
    );
}

#[test]
fn match_list_length_two() {
    assert_eq!(
        run_python_one("xs = [10, 20]\nmatch xs:\n case [p, q]:\n  print(p, q)\n"),
        "10 20"
    );
}

#[test]
fn match_list_head_tail() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nmatch xs:\n case [h, *t]:\n  print(h, len(t))\n"),
        "1 2"
    );
}

#[test]
fn match_empty_list() {
    assert_eq!(
        run_python_one("xs = []\nmatch xs:\n case []:\n  print('empty')\n"),
        "empty"
    );
}

#[test]
fn match_dict_two_keys() {
    assert_eq!(
        run_python_one(
            "d = {'a': 1, 'b': 2}\nmatch d:\n case {'a': av, 'b': bv}:\n  print(av + bv)\n"
        ),
        "3"
    );
}

#[test]
fn match_dict_rest_pattern() {
    assert_eq!(
        run_python_one(
            "d = {'a': 1, 'b': 2}\nmatch d:\n case {'a': av, **rest}:\n  print(av, 'b' in rest)\n"
        ),
        "1 True"
    );
}

#[test]
fn match_int_and_str_or() {
    assert_eq!(
        run_python_one("x = 'x'\nmatch x:\n case 1 | 'x':\n  print('hit')\n"),
        "hit"
    );
}

#[test]
fn match_guard_false_falls_through() {
    assert_eq!(
        run_python_one(
            "x = 1\nmatch x:\n case n if n > 5:\n  print('big')\n case _:\n  print('small')\n"
        ),
        "small"
    );
}

#[test]
fn match_on_enum_like_strings() {
    assert_eq!(
        run_python_one(
            "state = 'ready'\nmatch state:\n case 'ready' | 'running':\n  print('ok')\n"
        ),
        "ok"
    );
}

#[test]
fn match_float_literal() {
    assert_eq!(
        run_python_one("x = 1.5\nmatch x:\n case 1.5:\n  print('f')\n"),
        "f"
    );
}

#[test]
fn match_negative_int() {
    assert_eq!(
        run_python_one("x = -1\nmatch x:\n case -1:\n  print('neg')\n"),
        "neg"
    );
}

#[test]
fn match_in_function() {
    assert_eq!(
        run_python_one(
            "def label(n):\n match n:\n  case 0:\n   return 'zero'\n  case _:\n   return 'other'\nprint(label(0))\n"
        ),
        "zero"
    );
}

#[test]
fn match_with_return_in_case() {
    assert_eq!(
        run_python_one(
            "def f(x):\n match x:\n  case 1:\n   return 'a'\n  case _:\n   return 'b'\nprint(f(2))\n"
        ),
        "b"
    );
}

#[test]
fn match_bytes_pattern() {
    assert_eq!(
        run_python_one("b = b'ab'\nmatch b:\n case b'ab':\n  print('bytes')\n"),
        "bytes"
    );
}

#[test]
fn match_singleton_tuple_pattern() {
    assert_eq!(
        run_python_one("t = (7,)\nmatch t:\n case (v,):\n  print(v)\n"),
        "7"
    );
}

#[test]
fn match_star_starts_with_zero() {
    assert_eq!(
        run_python_one("xs = [0, 1, 2]\nmatch xs:\n case [0, *rest]:\n  print(rest)\n"),
        "[1, 2]"
    );
}

#[test]
fn match_multiple_types_wildcard() {
    assert_eq!(
        run_python_one(
            "def show(x):\n match x:\n  case str():\n   return 's'\n  case int():\n   return 'i'\n  case _:\n   return 'o'\nprint(show(3), show('a'))\n"
        ),
        "i s"
    );
}

#[test]
fn match_class_attr_pattern() {
    assert_eq!(
        run_python_one(
            "class P:\n def __init__(self, x):\n  self.x = x\np = P(4)\nmatch p:\n case P(x=x) if x > 0:\n  print(x)\n"
        ),
        "4"
    );
}

#[test]
fn match_nested_match() {
    assert_eq!(
        run_python_one(
            "x = (1, 2)\nmatch x:\n case (a, b):\n  match b:\n   case 2:\n    print('two')\n"
        ),
        "two"
    );
}

#[test]
fn match_list_single_element() {
    assert_eq!(
        run_python_one("xs = [9]\nmatch xs:\n case [n]:\n  print(n)\n"),
        "9"
    );
}

#[test]
fn match_set_display_not_used_match_int() {
    assert_eq!(
        run_python_one("x = {1, 2, 3}\nmatch len(x):\n case 3:\n  print('three')\n"),
        "three"
    );
}

#[test]
fn match_on_len_of_string() {
    assert_eq!(
        run_python_one("s = 'abc'\nmatch len(s):\n case 3:\n  print('len3')\n"),
        "len3"
    );
}

#[test]
fn match_on_type_via_wildcard() {
    assert_eq!(
        run_python_one("x = []\nmatch x:\n case list():\n  print('list')\n"),
        "list"
    );
}

#[test]
fn match_on_type_dict() {
    assert_eq!(
        run_python_one("x = {}\nmatch x:\n case dict():\n  print('dict')\n"),
        "dict"
    );
}

#[test]
fn match_complex_or_and_guard() {
    assert_eq!(
        run_python_one("x = 4\nmatch x:\n case 2 | 4 if x % 2 == 0:\n  print('even')\n"),
        "even"
    );
}

#[test]
fn match_assign_case_body_var() {
    assert_eq!(
        run_python_one("x = [1, 2]\nmatch x:\n case [a, b]:\n  s = a + b\nprint(s)\n"),
        "3"
    );
}

#[test]
fn match_break_in_loop() {
    assert_eq!(
        run_python_one("for v in [1, 2]:\n match v:\n  case 2:\n   print('stop')\n   break\n"),
        "stop"
    );
}

#[test]
fn match_continue_in_loop() {
    assert_eq!(
        run_python_one(
            "out = []\nfor v in [1, 2, 3]:\n match v:\n  case 2:\n   continue\n out.append(v)\nprint(out)\n"
        ),
        "[1, 3]"
    );
}

#[test]
fn match_tuple_of_lists() {
    assert_eq!(
        run_python_one("t = ([1], [2])\nmatch t:\n case ([a], [b]):\n  print(a, b)\n"),
        "1 2"
    );
}

#[test]
fn match_string_prefix_style() {
    assert_eq!(
        run_python_one(
            "s = 'error: msg'\nmatch s:\n case str() if s.startswith('error'):\n  print('err')\n"
        ),
        "err"
    );
}

#[test]
fn match_on_boolean_and() {
    assert_eq!(
        run_python_one("ok = True\nmatch ok:\n case True:\n  print('yes')\n"),
        "yes"
    );
}
