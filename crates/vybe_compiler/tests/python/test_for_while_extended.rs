crate::runtime_case!(
    for_range_basic,
    "for i in range(3):\n print(i)\n break\n",
    "0"
);
crate::runtime_case!(
    for_else_runs,
    "for i in range(2):\n pass\nelse:\n print('else')\n",
    "else"
);
crate::runtime_case!(
    for_else_break_skips,
    "for i in range(3):\n if i == 1:\n  break\nelse:\n print('else')\nprint('done')\n",
    "done"
);
crate::runtime_case!(
    while_else_runs,
    "i = 0\nwhile i < 2:\n i += 1\nelse:\n print('else')\n",
    "else"
);
crate::runtime_case!(
    while_else_break_skips,
    "i = 0\nwhile i < 3:\n if i == 1:\n  break\n i += 1\nelse:\n print('else')\nprint('fin')\n",
    "fin"
);
crate::runtime_case!(
    for_break_inner,
    "for i in range(3):\n for j in range(3):\n  if j == 1:\n   break\n print(i, j)\n break\n",
    "0 0"
);
crate::runtime_case!(
    for_continue_skip,
    "for i in range(4):\n if i % 2 == 0:\n  continue\n print(i)\n break\n",
    "1"
);
crate::runtime_case!(
    while_continue,
    "i = 0\nwhile i < 4:\n i += 1\n if i % 2 == 0:\n  continue\n print(i)\n break\n",
    "1"
);
crate::runtime_case!(
    for_enumerate,
    "for i, v in enumerate(['a', 'b']):\n print(i, v)\n break\n",
    "0 a"
);
crate::runtime_case!(
    for_zip,
    "for a, b in zip([1, 2], [3, 4]):\n print(a + b)\n break\n",
    "4"
);
crate::runtime_case!(
    for_dict_items,
    "for k, v in {'x': 1}.items():\n print(k, v)\n",
    "x 1"
);
crate::runtime_case!(
    for_dict_keys,
    "for k in {'a': 1, 'b': 2}:\n print(k)\n break\n",
    "a"
);
crate::runtime_case!(
    for_list_mutate,
    "a = [1, 2, 3]\nfor i, v in enumerate(a):\n a[i] = v * 2\nprint(a)\n",
    "[2, 4, 6]"
);
crate::runtime_case!(
    for_nested_sum,
    "s = 0\nfor i in range(2):\n for j in range(2):\n  s += i + j\nprint(s)\n",
    "4"
);
crate::runtime_case!(
    while_counter,
    "n = 0\ns = 0\nwhile n < 5:\n s += n\n n += 1\nprint(s)\n",
    "10"
);
crate::runtime_case!(
    while_true_break,
    "while True:\n print('once')\n break\n",
    "once"
);
crate::runtime_case!(
    for_string_chars,
    "s = ''\nfor c in 'ab':\n s += c\nprint(s)\n",
    "ab"
);
crate::runtime_case!(
    for_tuple_unpack,
    "for a, b in [(1, 2)]:\n print(a + b)\n",
    "3"
);
crate::runtime_case!(
    for_star_unpack,
    "for h, *t in [(1, 2, 3)]:\n print(h, len(t))\n",
    "1 2"
);
crate::runtime_case!(
    for_reversed,
    "for x in reversed([1, 2, 3]):\n print(x)\n break\n",
    "3"
);
crate::runtime_case!(
    for_sorted,
    "for x in sorted([3, 1, 2]):\n print(x)\n break\n",
    "1"
);
crate::runtime_case!(
    for_set_iter,
    "for x in {1, 2}:\n print(x in {1, 2})\n break\n",
    "True"
);
crate::runtime_case!(
    for_generator,
    "for x in (i for i in range(2)):\n print(x)\n break\n",
    "0"
);
crate::runtime_case!(
    while_walrus,
    "data = [1, 2, 3]\ni = 0\nwhile (v := data[i] if i < len(data) else None) is not None:\n print(v)\n i += 1\n if i > 0:\n  break\n",
    "1"
);
crate::runtime_case!(
    for_else_return_in_body,
    "def f():\n for i in range(1):\n  return i\n else:\n  return -1\nprint(f())\n",
    "0"
);
crate::runtime_case!(
    while_nested_break,
    "for i in range(2):\n j = 0\n while j < 3:\n  if j == 1:\n   break\n  j += 1\n print(i, j)\n break\n",
    "0 1"
);
crate::runtime_case!(
    for_range_step,
    "for i in range(0, 6, 2):\n print(i)\n break\n",
    "0"
);
crate::runtime_case!(
    for_range_negative_step,
    "for i in range(3, 0, -1):\n print(i)\n break\n",
    "3"
);
crate::runtime_case!(
    while_decrement,
    "n = 3\nwhile n > 0:\n n -= 1\nprint(n)\n",
    "0"
);
crate::runtime_case!(
    for_empty_iterable,
    "for x in []:\n print('no')\nelse:\n print('empty')\n",
    "empty"
);
crate::runtime_case!(
    while_false_never,
    "while 0:\n print('no')\nelse:\n print('zero')\n",
    "zero"
);
crate::runtime_case!(
    for_list_comp_side_effect,
    "out = []\nfor x in range(3):\n out.append(x)\nprint(out)\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    for_break_then_else,
    "for i in range(1):\n break\nelse:\n print('else')\nprint('after')\n",
    "after"
);
crate::runtime_case!(
    while_complex_condition,
    "i = 0\nwhile i < 3 and i != 5:\n i += 1\nprint(i)\n",
    "3"
);
crate::runtime_case!(
    for_match_inside,
    "for x in [1, 2]:\n match x:\n  case 1:\n   print('one')\n   break\n",
    "one"
);
crate::runtime_case!(
    for_try_inside,
    "for x in [1]:\n try:\n  print(x)\n except:\n  pass\n",
    "1"
);
crate::runtime_case!(
    for_with_inside,
    "class CM:\n def __enter__(self):\n  return 1\n def __exit__(self, *a):\n  pass\nfor _ in range(1):\n with CM() as v:\n  print(v)\n",
    "1"
);
crate::runtime_case!(
    while_read_lines_mock,
    "lines = ['a', 'b']\ni = 0\nwhile i < len(lines):\n print(lines[i])\n i += 1\n break\n",
    "a"
);
crate::runtime_case!(
    for_accumulate_product,
    "p = 1\nfor x in [2, 3, 4]:\n p *= x\nprint(p)\n",
    "24"
);
crate::runtime_case!(
    for_filter_manual,
    "for x in [1, 2, 3]:\n if x < 2:\n  print(x)\n",
    "1"
);
crate::runtime_case!(
    for_early_return,
    "def f():\n for i in range(5):\n  if i == 2:\n   return i\n return -1\nprint(f())\n",
    "2"
);
crate::runtime_case!(
    while_flag_pattern,
    "done = False\nn = 0\nwhile not done:\n n += 1\n if n >= 2:\n  done = True\nprint(n)\n",
    "2"
);
crate::runtime_case!(
    for_slice_iter,
    "for x in [1, 2, 3][1:]:\n print(x)\n break\n",
    "2"
);
crate::runtime_case!(
    for_dict_values,
    "for v in {'a': 10}.values():\n print(v)\n",
    "10"
);
crate::runtime_case!(
    for_dict_items_unpack,
    "for k, v in [('a', 1)]:\n print(k, v)\n",
    "a 1"
);
crate::runtime_case!(for_bytes_iter, "for b in b'ab':\n print(b)\n break\n", "97");
crate::runtime_case!(
    while_else_continue_path,
    "i = 0\nwhile i < 2:\n i += 1\n continue\nelse:\n print('else')\n",
    "else"
);
crate::runtime_case!(
    for_nested_else,
    "for i in range(1):\n for j in range(1):\n  pass\n else:\n  print('inner')\n",
    "inner"
);

crate::compile_case!(
    for_async_iter,
    "async def f():\n async for x in agen():\n  pass\n"
);
crate::compile_case!(
    while_async,
    "async def f():\n while True:\n  await asyncio.sleep(0)\n  break\n"
);
crate::compile_case!(for_star_target, "for *a, b in [(1, 2, 3)]:\n pass\n");
crate::compile_case!(
    for_try_break,
    "for i in range(3):\n try:\n  break\n finally:\n  pass\n"
);
crate::compile_case!(
    while_try_continue,
    "i = 0\nwhile i < 2:\n i += 1\n try:\n  continue\n finally:\n  pass\n"
);
