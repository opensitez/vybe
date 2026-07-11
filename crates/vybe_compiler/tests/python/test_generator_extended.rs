//! Generator send/throw/close, yield from, and iterator protocol edge cases.

crate::runtime_case!(
    generator_yield_values,
    "def g():\n yield 1\n yield 2\nprint(list(g()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    generator_yield_from_list,
    "def g():\n yield from [1, 2, 3]\nprint(list(g()))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    generator_yield_from_range,
    "def g():\n yield from range(3)\nprint(list(g()))\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    generator_send_value,
    "def g():\n x = yield 1\n yield x\nit = g()\nprint(next(it))\nprint(it.send(9))\n",
    "1\n9"
);
crate::runtime_case!(
    generator_close_raises,
    "def g():\n try:\n  yield 1\n finally:\n  print('fin')\nit = g()\nprint(next(it))\nit.close()\n",
    "1\nfin"
);
crate::runtime_case!(
    generator_throw_caught,
    "def g():\n try:\n  yield 1\n except ValueError:\n  yield 2\nit = g()\nprint(next(it))\nprint(it.throw(ValueError))\n",
    "1\n2"
);
crate::runtime_case!(
    generator_return_value,
    "def g():\n yield 1\n return 99\nit = g()\nprint(list(it))\n",
    "[1]"
);
crate::runtime_case!(
    generator_expr_basic,
    "print(list(x * 2 for x in range(3)))\n",
    "[0, 2, 4]"
);
crate::runtime_case!(
    generator_expr_filter,
    "print(list(x for x in range(5) if x % 2))\n",
    "[1, 3]"
);
crate::runtime_case!(
    generator_iter_next,
    "def g():\n yield 'a'\n yield 'b'\nit = iter(g())\nprint(next(it), next(it))\n",
    "a b"
);
crate::runtime_case!(
    generator_stop_iteration,
    "def g():\n yield 1\ntry:\n next(g())\n next(g())\nexcept StopIteration:\n print('stop')\n",
    "stop"
);
crate::runtime_case!(
    generator_yield_in_loop,
    "def g():\n for i in range(3):\n  yield i\nprint(list(g()))\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    generator_nested_yield_from,
    "def inner():\n yield 2\ndef outer():\n yield 1\n yield from inner()\n yield 3\nprint(list(outer()))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    generator_delegation_return,
    "def inner():\n yield 1\n return 'done'\ndef outer():\n y = yield from inner()\n yield y\nprint(list(outer()))\n",
    "[1, 'done']"
);
crate::runtime_case!(
    generator_frame_locals,
    "def g():\n x = 10\n yield x\n x = 20\n yield x\nprint(list(g()))\n",
    "[10, 20]"
);
crate::runtime_case!(
    generator_param_binding,
    "def g(n):\n for i in range(n):\n  yield i\nprint(list(g(3)))\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    generator_multiple_yield_from,
    "def g():\n yield from [1]\n yield from [2]\nprint(list(g()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    generator_bool,
    "def g():\n yield 1\nprint(bool(g()))\n",
    "True"
);
crate::runtime_case!(
    generator_iter_idempotent,
    "def g():\n yield 1\nit = g()\nprint(list(it) == list(g()))\n",
    "True"
);
crate::runtime_case!(
    generator_comp_nested,
    "print([[y for y in range(x)] for x in range(3)])\n",
    "[[], [0], [0, 1]]"
);
crate::runtime_case!(
    generator_enumerate_yield,
    "def g():\n for i, v in enumerate(['a', 'b']):\n  yield i, v\nprint(list(g()))\n",
    "[(0, 'a'), (1, 'b')]"
);
crate::runtime_case!(
    generator_zip_yield,
    "def g():\n for a, b in zip([1, 2], [3, 4]):\n  yield a + b\nprint(list(g()))\n",
    "[4, 6]"
);
crate::runtime_case!(
    generator_infinite_take,
    "def count():\n n = 0\n while True:\n  yield n\n  n += 1\nprint([next(count()) for _ in range(3)])\n",
    "[0, 0, 0]"
);
crate::runtime_case!(
    generator_reentrant_same,
    "def g():\n yield 1\ng1 = g()\nprint(next(g1))\nprint(next(g()))\n",
    "1\n1"
);
crate::runtime_case!(
    generator_yield_none_explicit,
    "def g():\n yield None\nprint(list(g()))\n",
    "[None]"
);
crate::runtime_case!(
    generator_expr_map,
    "print(list(map(lambda x: x + 1, (i for i in range(3)))))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    generator_filter,
    "print(list(filter(None, (0, 1, 0, 2))))\n",
    "[1, 2]"
);
crate::runtime_case!(generator_sum, "print(sum(i for i in range(5)))\n", "10");
crate::runtime_case!(
    generator_any,
    "print(any(i > 3 for i in range(5)))\n",
    "True"
);
crate::runtime_case!(
    generator_all,
    "print(all(i >= 0 for i in range(3)))\n",
    "True"
);
crate::runtime_case!(generator_max, "print(max(i for i in [3, 1, 2]))\n", "3");
crate::runtime_case!(generator_min, "print(min(i for i in [3, 1, 2]))\n", "1");
crate::runtime_case!(
    generator_sorted,
    "print(sorted((i for i in [3, 1, 2])))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    generator_list_ctor,
    "print(list(i for i in range(2)))\n",
    "[0, 1]"
);
crate::runtime_case!(
    generator_tuple_ctor,
    "print(tuple(i for i in range(2)))\n",
    "(0, 1)"
);
crate::runtime_case!(
    generator_set_ctor,
    "print(sorted(i for i in [2, 1, 2]))\n",
    "[1, 2]"
);
crate::runtime_case!(
    generator_dict_comp_not_gen,
    "print({i: i for i in range(2)})\n",
    "{0: 0, 1: 1}"
);
crate::runtime_case!(
    generator_yield_string_join,
    "print(''.join(ch for ch in 'ab'))\n",
    "ab"
);
crate::runtime_case!(
    generator_chunk_pattern,
    "def chunks(it, n):\n buf = []\n for x in it:\n  buf.append(x)\n  if len(buf) == n:\n   yield buf\n   buf = []\nprint(list(chunks([1,2,3,4], 2)))\n",
    "[[1, 2], [3, 4]]"
);
crate::runtime_case!(
    generator_tee_manual,
    "def g():\n yield 1\n yield 2\na = list(g())\nb = list(g())\nprint(a == b)\n",
    "True"
);
crate::runtime_case!(
    generator_throw_exit,
    "def g():\n yield 1\ntry:\n g().throw(TypeError)\nexcept TypeError:\n print('typed')\n",
    "typed"
);
crate::runtime_case!(
    generator_finally_yield,
    "def g():\n try:\n  yield 1\n finally:\n  pass\nprint(list(g()))\n",
    "[1]"
);
crate::runtime_case!(
    generator_yield_from_empty,
    "def g():\n yield from []\n yield 1\nprint(list(g()))\n",
    "[1]"
);
crate::runtime_case!(
    generator_mutable_state,
    "def g():\n acc = []\n for i in range(3):\n  acc.append(i)\n  yield sum(acc)\nprint(list(g()))\n",
    "[0, 1, 3]"
);
crate::runtime_case!(
    generator_break_inside,
    "def g():\n for i in range(10):\n  if i == 3:\n   break\n  yield i\nprint(list(g()))\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    generator_continue_inside,
    "def g():\n for i in range(4):\n  if i % 2 == 0:\n   continue\n  yield i\nprint(list(g()))\n",
    "[1, 3]"
);

crate::compile_case!(
    generator_yield_from_generator,
    "def inner():\n yield 1\ndef outer():\n yield from inner()\nlist(outer())\n"
);
crate::compile_case!(generator_async_def, "async def ag():\n yield 1\n");
crate::compile_case!(generator_peek_pattern, "def g():\n yield 1\nit = g()\n");
crate::compile_case!(
    generator_throw_while_running,
    "def g():\n yield 1\n yield 2\nit = g()\nnext(it)\nit.throw(RuntimeError)\n"
);
crate::compile_case!(
    generator_yield_classmethod,
    "class C:\n @classmethod\n def f(cls):\n  yield cls\nlist(C.f())\n"
);
