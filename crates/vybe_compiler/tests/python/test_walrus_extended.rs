//! Walrus operator in comprehensions, loops, if, and nested scopes.


crate::runtime_case!(
    walrus_if_body,
    "x = 0\nif (n := 5) > 3:\n x = n\nprint(x)\n",
    "5"
);
crate::runtime_case!(
    walrus_while_read,
    "s = 'abc'\ni = 0\nwhile (c := s[i] if i < len(s) else ''):\n i += 1\n if i >= 2:\n  break\nprint(c)\n",
    "b"
);
crate::runtime_case!(
    walrus_list_comp_filter,
    "print([y for x in [1, 2, 3] if (y := x * 2) > 2])\n",
    "[4, 6]"
);
crate::runtime_case!(
    walrus_dict_comp,
    "print({k: (v := k * 2) for k in range(3)})\n",
    "{0: 0, 1: 2, 2: 4}"
);
crate::runtime_case!(
    walrus_nested_paren,
    "print((a := 1) + (b := 2))\n",
    "3"
);
crate::runtime_case!(
    walrus_assign_then_use,
    "if (x := 10):\n print(x)\n",
    "10"
);
crate::runtime_case!(
    walrus_false_branch_skipped,
    "y = 1\nif (z := 0):\n y = 9\nprint(y)\n",
    "1"
);
crate::runtime_case!(
    walrus_in_expression,
    "print([(a := i) for i in range(2)][-1])\n",
    "1"
);
crate::runtime_case!(
    walrus_string_length,
    "s = 'hello'\nprint(len(s) if (n := len(s)) else 0)\n",
    "5"
);
crate::runtime_case!(
    walrus_list_append_in_loop,
    "out = []\nfor i in range(3):\n if (v := i * i) >= 0:\n  out.append(v)\nprint(out)\n",
    "[0, 1, 4]"
);
crate::runtime_case!(
    walrus_comprehension_value,
    "print([ (x := i + 1) for i in range(2) ])\n",
    "[1, 2]"
);
crate::runtime_case!(
    walrus_or_short_circuit,
    "print((a := 0) or (b := 5))\n",
    "5"
);
crate::runtime_case!(
    walrus_and_short_circuit,
    "print((a := 3) and (b := 4))\n",
    "4"
);
crate::runtime_case!(
    walrus_in_fstring,
    "x = 7\nprint(f'{(y := x + 1)}')\n",
    "8"
);
crate::runtime_case!(
    walrus_function_arg,
    "def f(v):\n return v\nprint(f((z := 9)))\n",
    "9"
);
crate::runtime_case!(
    walrus_tuple_unpack,
    "(a := 1)\n(b := 2)\nprint(a, b)\n",
    "1 2"
);
crate::runtime_case!(
    walrus_set_comp,
    "print({(v := x % 2) for x in range(4)})\n",
    "{0, 1}"
);
crate::runtime_case!(
    walrus_gen_exp,
    "print(list((y := x) for x in range(2)))\n",
    "[0, 1]"
);
crate::runtime_case!(
    walrus_chained_compare,
    "print(1 < (n := 2) < 3)\n",
    "True"
);
crate::runtime_case!(
    walrus_attribute_read,
    "class C:\n x = 5\nc = C()\nprint((v := c.x))\n",
    "5"
);
crate::runtime_case!(
    walrus_subscript_read,
    "d = {'k': 9}\nprint((v := d['k']))\n",
    "9"
);
crate::runtime_case!(
    walrus_slice_len,
    "a = [1, 2, 3]\nprint(len(b) if (b := a[:2]) else 0)\n",
    "2"
);
crate::runtime_case!(
    walrus_for_else,
    "for i in range(1):\n if (x := i) == 0:\n  print('ok')\nelse:\n print('skip')\n",
    "ok"
);
crate::runtime_case!(
    walrus_try_except,
    "try:\n raise ValueError('e')\nexcept ValueError as e:\n if (msg := str(e)):\n  print(len(msg))\n",
    "1"
);
crate::runtime_case!(
    walrus_lambda_body,
    "f = lambda: (n := 3)\nprint(f())\n",
    "3"
);
crate::runtime_case!(
    walrus_list_index,
    "a = [10, 20]\nprint(a[(i := 1)])\n",
    "20"
);
crate::runtime_case!(
    walrus_bool_context,
    "print(bool((x := 0)))\n",
    "False"
);
crate::runtime_case!(
    walrus_truthy_context,
    "print(bool((x := 1)))\n",
    "True"
);
crate::runtime_case!(
    walrus_while_counter,
    "n = 3\nc = 0\nwhile (n := n - 1) >= 0:\n c += 1\nprint(c)\n",
    "4"
);
crate::runtime_case!(
    walrus_match_guard,
    "x = 5\nmatch x:\n case n if (d := n // 2) > 0:\n  print(d)\n",
    "2"
);
crate::runtime_case!(
    walrus_nested_comp,
    "print([ (a := i) + (b := 1) for i in range(2) ])\n",
    "[1, 2]"
);
crate::runtime_case!(
    walrus_del_after_assign,
    "if (x := [1, 2]):\n del x[0]\nprint('done')\n",
    "done"
);
crate::runtime_case!(
    walrus_augmented,
    "x = 1\n(x := x + 2)\nprint(x)\n",
    "3"
);
crate::runtime_case!(
    walrus_in_dict_get,
    "d = {'a': 1}\nprint(d.get('a') if (k := 'a') else 0)\n",
    "1"
);
crate::runtime_case!(
    walrus_enumerate,
    "print([ (i := idx) for idx, _ in enumerate(['x', 'y']) ])\n",
    "[0, 1]"
);
crate::runtime_case!(
    walrus_zip_unpack,
    "print([ (a := x) + (b := y) for x, y in zip([1, 2], [10, 20]) ])\n",
    "[11, 22]"
);
crate::runtime_case!(
    walrus_any_all,
    "print(any((v := x) > 2 for x in [1, 3]))\n",
    "True"
);
crate::runtime_case!(
    walrus_sorted_key,
    "print(sorted(['bb', 'a'], key=lambda s: (n := len(s))))\n",
    "['a', 'bb']"
);
crate::runtime_case!(
    walrus_re_match,
    "import re\nm = re.match('(a+)', 'aaa')\nprint(m.group(1) if (m := re.match('(a+)', 'aaa')) else '')\n",
    "aaa"
);
crate::runtime_case!(
    walrus_type_check,
    "print(type((n := 42)).__name__)\n",
    "int"
);
crate::runtime_case!(
    walrus_isinstance,
    "print(isinstance((x := 'hi'), str))\n",
    "True"
);
crate::runtime_case!(
    walrus_len_in_cond,
    "s = 'abc'\nprint('yes' if (n := len(s)) == 3 else 'no')\n",
    "yes"
);
crate::runtime_case!(
    walrus_min_max,
    "print(max((v := x) for x in [1, 5, 3]))\n",
    "5"
);
crate::runtime_case!(
    walrus_sum_comp,
    "print(sum((v := i) for i in range(4)))\n",
    "6"
);

crate::compile_case!(walrus_named_expr_in_lambda, "f = lambda x: (y := x)\n");
crate::compile_case!(walrus_in_class_body, "class C:\n x = (y := 1)\n");
crate::compile_case!(walrus_with_statement, "class CM:\n def __enter__(self): return self\n def __exit__(self, *a): pass\nwith CM() as (c := CM()):\n pass\n");
crate::compile_case!(walrus_async_comp, "async def f():\n return [(x := i) async for i in async_range(2)]\n");
crate::compile_case!(walrus_yield_from, "def g():\n if (x := 1):\n  yield x\n");
