//! Extended match/case: guards, class patterns, mapping patterns, or-patterns, as-names.

use crate::helpers::*;

crate::runtime_case!(
    match_int_literal,
    "x = 42\nmatch x:\n case 42:\n  print('yes')\n",
    "yes"
);
crate::runtime_case!(
    match_str_literal,
    "x = 'ok'\nmatch x:\n case 'ok':\n  print(1)\n",
    "1"
);
crate::runtime_case!(
    match_wildcard,
    "x = object()\nmatch x:\n case _:\n  print('any')\n",
    "any"
);
crate::runtime_case!(
    match_or_pattern,
    "x = 2\nmatch x:\n case 1 | 2 | 3:\n  print('small')\n",
    "small"
);
crate::runtime_case!(
    match_list_destructure,
    "x = [1, 2, 3]\nmatch x:\n case [a, b, c]:\n  print(a + c)\n",
    "4"
);
crate::runtime_case!(
    match_list_star_tail,
    "x = [1, 2, 3, 4]\nmatch x:\n case [h, *t]:\n  print(h, len(t))\n",
    "1 3"
);
crate::runtime_case!(
    match_list_empty,
    "x = []\nmatch x:\n case []:\n  print('empty')\n",
    "empty"
);
crate::runtime_case!(
    match_tuple_destructure,
    "x = (10, 20)\nmatch x:\n case (a, b):\n  print(a * b)\n",
    "200"
);
crate::runtime_case!(
    match_nested_tuple,
    "x = (1, (2, 3))\nmatch x:\n case (a, (b, c)):\n  print(b + c)\n",
    "5"
);
crate::runtime_case!(
    match_dict_exact,
    "d = {'x': 1, 'y': 2}\nmatch d:\n case {'x': xv, 'y': yv}:\n  print(xv + yv)\n",
    "3"
);
crate::runtime_case!(
    match_dict_capture_rest,
    "d = {'a': 1, 'b': 2}\nmatch d:\n case {'a': av, **rest}:\n  print(av, 'b' in rest)\n",
    "1 True"
);
crate::runtime_case!(
    match_guard_true,
    "x = 5\nmatch x:\n case n if n > 3:\n  print('big')\n",
    "big"
);
crate::runtime_case!(
    match_guard_false_fallthrough,
    "x = 2\nmatch x:\n case n if n > 3:\n  print('big')\n case _:\n  print('other')\n",
    "other"
);
crate::runtime_case!(
    match_as_name,
    "x = 7\nmatch x:\n case n as value:\n  print(value)\n",
    "7"
);
crate::runtime_case!(
    match_class_pattern,
    "class P:\n def __init__(self, x, y):\n  self.x = x\n  self.y = y\np = P(2, 3)\nmatch p:\n case P(x, y):\n  print(x, y)\n",
    "2 3"
);
crate::runtime_case!(
    match_class_attr_guard,
    "class P:\n def __init__(self, v):\n  self.v = v\np = P(9)\nmatch p:\n case P(v=v) if v > 5:\n  print(v)\n",
    "9"
);
crate::runtime_case!(
    match_true_singleton,
    "x = True\nmatch x:\n case True:\n  print('t')\n",
    "t"
);
crate::runtime_case!(
    match_none_singleton,
    "x = None\nmatch x:\n case None:\n  print('nil')\n",
    "nil"
);
crate::runtime_case!(
    match_bytes_literal,
    "x = b'ab'\nmatch x:\n case b'ab':\n  print('bytes')\n",
    "bytes"
);
crate::runtime_case!(
    match_first_branch_wins,
    "x = 1\nmatch x:\n case 1:\n  print('first')\n case 1:\n  print('second')\n",
    "first"
);
crate::runtime_case!(
    match_no_case_unmatched,
    "x = 99\nmatched = False\nmatch x:\n case 1:\n  matched = True\nprint(matched)\n",
    "False"
);
crate::runtime_case!(
    match_sequence_length_mismatch,
    "x = [1, 2]\nmatch x:\n case [a, b, c]:\n  print('no')\n case _:\n  print('fallback')\n",
    "fallback"
);
crate::runtime_case!(
    match_set_display,
    "x = {1, 2}\nmatch x:\n case {1, 2}:\n  print('set')\n case _:\n  print('no')\n",
    "set"
);
crate::runtime_case!(
    match_enum_like_int,
    "x = 0\nmatch x:\n case 0:\n  print('zero')\n",
    "zero"
);
crate::runtime_case!(
    match_float_literal,
    "x = 1.5\nmatch x:\n case 1.5:\n  print('f')\n",
    "f"
);
crate::runtime_case!(
    match_negative_int,
    "x = -1\nmatch x:\n case -1:\n  print('neg')\n",
    "neg"
);
crate::runtime_case!(
    match_list_middle_star,
    "x = [1, 2, 3, 4]\nmatch x:\n case [1, *mid, 4]:\n  print(len(mid))\n",
    "2"
);
crate::runtime_case!(
    match_tuple_mixed,
    "x = (1, [2, 3])\nmatch x:\n case (a, [b, c]):\n  print(a + b + c)\n",
    "6"
);
crate::runtime_case!(
    match_nested_match,
    "x = 2\nmatch x:\n case 2:\n  match x:\n   case 2:\n    print('nested')\n",
    "nested"
);
crate::runtime_case!(
    match_in_function,
    "def f(v):\n match v:\n  case 1:\n   return 'one'\n  case _:\n   return 'other'\nprint(f(1))\n",
    "one"
);
crate::runtime_case!(
    match_in_loop,
    "out = []\nfor v in [1, 2, 3]:\n match v:\n  case 2:\n   out.append('two')\nprint(out)\n",
    "['two']"
);
crate::runtime_case!(
    match_with_else,
    "x = 5\nmatch x:\n case 1:\n  print('one')\n case _:\n  print('done')\n",
    "done"
);
crate::runtime_case!(
    match_capture_subpattern,
    "x = [1, 2]\nmatch x:\n case [a, b] as pair:\n  print(len(pair))\n",
    "2"
);
crate::runtime_case!(
    match_mapping_single_key,
    "d = {'only': 1}\nmatch d:\n case {'only': v}:\n  print(v)\n",
    "1"
);
crate::runtime_case!(
    match_class_kw_pattern,
    "class P:\n def __init__(self, x):\n  self.x = x\np = P(4)\nmatch p:\n case P(x=4):\n  print('kw')\n",
    "kw"
);
crate::runtime_case!(
    match_or_with_guard,
    "x = 3\nmatch x:\n case 1 | 2:\n  print('low')\n case 3 | 4 if x == 3:\n  print('mid')\n",
    "mid"
);
crate::runtime_case!(
    match_list_first_only,
    "x = [9, 8, 7]\nmatch x:\n case [first, *_]:\n  print(first)\n",
    "9"
);
crate::runtime_case!(
    match_str_prefix,
    "s = 'hello'\nmatch s:\n case str() if s.startswith('he'):\n  print('prefix')\n",
    "prefix"
);
crate::runtime_case!(
    match_return_from_case,
    "def f():\n x = 1\n match x:\n  case 1:\n   return 99\n return 0\nprint(f())\n",
    "99"
);
crate::runtime_case!(
    match_break_in_loop,
    "for i in range(3):\n match i:\n  case 1:\n   print('hit')\n   break\n",
    "hit"
);
crate::runtime_case!(
    match_continue_in_loop,
    "for i in range(3):\n match i:\n  case 0:\n   continue\n print(i)\n",
    "1"
);
crate::runtime_case!(
    match_type_name,
    "x = 1\nmatch x:\n case int():\n  print('int')\n case _:\n  print('other')\n",
    "int"
);
crate::runtime_case!(
    match_singleton_bool_false,
    "x = False\nmatch x:\n case False:\n  print('f')\n",
    "f"
);
crate::runtime_case!(
    match_empty_dict,
    "x = {}\nmatch x:\n case {}:\n  print('empty')\n",
    "empty"
);
crate::runtime_case!(
    match_mixed_pos_kw_class,
    "class P:\n def __init__(self, a, b=0):\n  self.a = a\n  self.b = b\np = P(1, 2)\nmatch p:\n case P(1, b=2):\n  print('mix')\n",
    "mix"
);
crate::runtime_case!(
    match_sequence_unpack_star,
    "x = [1, 2, 3]\nmatch x:\n case [1, *rest]:\n  print(sum(rest))\n",
    "5"
);

crate::compile_case!(match_soft_keyword, "match x:\n case 1:\n  pass\n");
crate::compile_case!(match_capture_walrus, "match x:\n case n if (d := n // 2):\n  pass\n");
crate::compile_case!(match_class_nested, "class A:\n pass\nmatch o:\n case A():\n  pass\n");
crate::compile_case!(match_sequence_or, "match x:\n case [1, 2] | [3, 4]:\n  pass\n");
crate::compile_case!(match_mapping_or, "match d:\n case {'a': 1} | {'b': 2}:\n  pass\n");
