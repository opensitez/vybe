//! Chained comparisons, Ellipsis, NotImplemented, pprint/repr/ascii.

crate::runtime_case!(
    chained_lt_lt,
    "print(1 < 2 < 3)\n",
    "True"
);
crate::runtime_case!(
    chained_lt_lt_false,
    "print(1 < 3 < 2)\n",
    "False"
);
crate::runtime_case!(
    chained_eq_eq,
    "print(1 == 1 == 1)\n",
    "True"
);
crate::runtime_case!(
    chained_ne_ne,
    "print(1 != 2 != 3)\n",
    "True"
);
crate::runtime_case!(
    chained_le_le,
    "print(1 <= 2 <= 2)\n",
    "True"
);
crate::runtime_case!(
    chained_ge_ge,
    "print(3 >= 2 >= 1)\n",
    "True"
);
crate::runtime_case!(
    chained_mixed,
    "print(0 < 1 == 1)\n",
    "True"
);
crate::runtime_case!(
    chained_short_circuit,
    "def f():\n return 2\nprint(1 < f() < 3)\n",
    "True"
);
crate::runtime_case!(
    chained_with_expr,
    "x = 5\nprint(0 < x < 10)\n",
    "True"
);
crate::runtime_case!(
    chained_string_len,
    "s = 'abc'\nprint(1 < len(s) < 5)\n",
    "True"
);
crate::runtime_case!(
    ellipsis_singleton,
    "print(Ellipsis is ...)\n",
    "True"
);
crate::runtime_case!(
    ellipsis_type,
    "print(type(...).__name__)\n",
    "ellipsis"
);
crate::runtime_case!(
    ellipsis_in_slice,
    "print([1, 2, 3, 4][1:...])\n",
    "[2, 3, 4]"
);
crate::runtime_case!(
    ellipsis_tuple,
    "print((..., 1)[1])\n",
    "1"
);
crate::runtime_case!(
    notimplemented_singleton,
    "print(NotImplemented)\n",
    "NotImplemented"
);
crate::runtime_case!(
    notimplemented_bool,
    "print(bool(NotImplemented))\n",
    "True"
);
crate::runtime_case!(
    notimplemented_type,
    "print(type(NotImplemented).__name__)\n",
    "NotImplementedType"
);
crate::runtime_case!(
    richcompare_notimplemented,
    "class C:\n pass\nprint(C() == C())\n",
    "False"
);
crate::runtime_case!(
    repr_list,
    "print(repr([1, 2]))\n",
    "[1, 2]"
);
crate::runtime_case!(
    repr_dict,
    "print(repr({'a': 1}))\n",
    "{'a': 1}"
);
crate::runtime_case!(
    repr_str_escapes,
    "print(repr('\\n'))\n",
    "'\\n'"
);
crate::runtime_case!(
    ascii_str,
    "print(ascii('hi'))\n",
    "'hi'"
);
crate::runtime_case!(
    ascii_non_ascii,
    "print(ascii('é'))\n",
    "'\\xe9'"
);
crate::runtime_case!(
    pprint_list,
    "import pprint\nimport io\nbuf = io.StringIO()\npprint.pprint([1, 2, 3], stream=buf)\nprint('1' in buf.getvalue())\n",
    "True"
);
crate::runtime_case!(
    pprint_dict,
    "import pprint\nprint(isinstance(pprint.pformat({'a': 1}), str))\n",
    "True"
);
crate::runtime_case!(
    pprint_depth,
    "import pprint\nprint(pprint.pformat([[1]], depth=1))\n",
    "[[...]]"
);
crate::runtime_case!(
    pprint_compact,
    "import pprint\nprint(pprint.pformat([1, 2, 3], width=20))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    repr_bytes,
    "print(repr(b'hi'))\n",
    "b'hi'"
);
crate::runtime_case!(
    repr_set,
    "print(repr({1, 2}))\n",
    "{1, 2}"
);
crate::runtime_case!(
    repr_frozenset,
    "print(repr(frozenset({1})))\n",
    "frozenset({1})"
);
crate::runtime_case!(
    repr_range,
    "print(repr(range(3)))\n",
    "range(0, 3)"
);
crate::runtime_case!(
    repr_bool,
    "print(repr(True))\n",
    "True"
);
crate::runtime_case!(
    repr_none,
    "print(repr(None))\n",
    "None"
);
crate::runtime_case!(
    repr_float,
    "print(repr(1.5))\n",
    "1.5"
);
crate::runtime_case!(
    repr_int,
    "print(repr(42))\n",
    "42"
);
crate::runtime_case!(
    chained_in_membership,
    "print(2 in [1, 2, 3] in [True])\n",
    "True"
);
crate::runtime_case!(
    chained_is_identity,
    "a = []\nprint(a is a is a)\n",
    "True"
);
crate::runtime_case!(
    pprint_isreadable,
    "import pprint\nprint(pprint.isreadable([1, 2]))\n",
    "True"
);
crate::runtime_case!(
    pprint_isrecursive,
    "import pprint\na = []\na.append(a)\nprint(pprint.isrecursive(a))\n",
    "True"
);
crate::runtime_case!(
    pprint_pformat,
    "import pprint\nprint(pprint.pformat('x'))\n",
    "'x'"
);
crate::runtime_case!(
    pprint_pp_function,
    "import pprint\nprint(callable(pprint.pp))\n",
    "True"
);
crate::runtime_case!(
    repr_custom,
    "class C:\n def __repr__(self):\n  return 'C()'\nprint(repr(C()))\n",
    "C()"
);
crate::runtime_case!(
    ascii_custom,
    "class C:\n def __repr__(self):\n  return 'é'\nprint(ascii(C()))\n",
    "'\\xe9'"
);
crate::runtime_case!(
    chained_comparison_types,
    "print(1 < 2.0 < 3)\n",
    "True"
);
crate::runtime_case!(
    ellipsis_not_equal_one,
    "print(... != 1)\n",
    "True"
);
crate::runtime_case!(
    notimplemented_not_equal,
    "print(NotImplemented != 1)\n",
    "True"
);

crate::compile_case!(chained_comparison_assign, "a = 1 < 2 < 3\n");
crate::compile_case!(ellipsis_annotation, "def f(x: ...): pass\n");
crate::compile_case!(notimplemented_return, "def f():\n return NotImplemented\n");
crate::compile_case!(pprint_saferepr, "import pprint\npprint.saferepr([1])\n");
crate::compile_case!(repr_recursive, "a = []\na.append(a)\nrepr(a)\n");
