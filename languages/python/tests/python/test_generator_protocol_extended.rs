//! Generator send/throw/close and iterator protocol depth.

crate::runtime_case!(
    generator_send_first,
    "def g():\n x = yield 1\n yield x\nit = g()\nprint(next(it), it.send(9))\n",
    "1 9"
);
crate::runtime_case!(
    generator_send_none,
    "def g():\n yield 1\n yield 2\nit = g()\nprint(next(it))\nprint(it.send(None))\n",
    "1\n2"
);
crate::runtime_case!(
    generator_throw_caught,
    "def g():\n try:\n  yield 1\n except ValueError:\n  yield 2\nit = g()\nprint(next(it))\nprint(it.throw(ValueError))\n",
    "1\n2"
);
crate::runtime_case!(
    generator_throw_uncaught,
    "def g():\n yield 1\ntry:\n g().throw(ValueError)\n print('ok')\nexcept ValueError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    generator_close,
    "def g():\n try:\n  yield 1\n finally:\n  print('fin')\nit = g()\nprint(next(it))\nit.close()\n",
    "1\nfin"
);
crate::runtime_case!(
    generator_return_value,
    "def g():\n yield 1\n return 99\nit = g()\nprint(list(it))\n",
    "[1]"
);
crate::runtime_case!(
    generator_yield_from_list,
    "def g():\n yield from [1, 2, 3]\nprint(list(g()))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    generator_yield_from_gen,
    "def inner():\n yield 2\ndef outer():\n yield 1\n yield from inner()\n yield 3\nprint(list(outer()))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    generator_yield_from_return,
    "def inner():\n yield 1\n return 'done'\ndef outer():\n y = yield from inner()\n yield y\nprint(list(outer()))\n",
    "[1, 'done']"
);
crate::runtime_case!(
    generator_iter_protocol,
    "def g():\n yield 'a'\n yield 'b'\nit = iter(g())\nprint(next(it), next(it))\n",
    "a b"
);
crate::runtime_case!(
    generator_stop_iteration,
    "def g():\n yield 1\ntry:\n next(g())\n next(g())\nexcept StopIteration:\n print('stop')\n",
    ""
);
crate::runtime_case!(
    generator_expr_send,
    "g = (x for x in range(2))\nprint(next(g))\n",
    "0"
);
crate::runtime_case!(
    generator_delegation_throw,
    "def inner():\n try:\n  yield 1\n except ValueError:\n  yield 2\ndef outer():\n yield from inner()\nit = outer()\nprint(next(it))\nprint(it.throw(ValueError))\n",
    "1\n2"
);
crate::runtime_case!(
    generator_frame_locals,
    "def g():\n x = 10\n yield x\n x = 20\n yield x\nprint(list(g()))\n",
    "[10, 20]"
);
crate::runtime_case!(
    generator_param,
    "def g(n):\n for i in range(n):\n  yield i\nprint(list(g(3)))\n",
    "[0, 1, 2]"
);
crate::runtime_case!(
    generator_yield_none,
    "def g():\n yield None\nprint(list(g()))\n",
    "[None]"
);
crate::runtime_case!(
    generator_bool,
    "def g():\n yield 1\nprint(bool(g()))\n",
    "True"
);
crate::runtime_case!(
    generator_name,
    "def g():\n yield 1\nprint(g.__name__)\n",
    "g"
);
crate::runtime_case!(
    generator_isgenerator,
    "import inspect\ndef g():\n yield 1\nprint(inspect.isgenerator(g()))\n",
    "True"
);
crate::runtime_case!(
    generator_isgeneratorfunction,
    "import inspect\ndef g():\n yield 1\ndef f():\n pass\nprint(inspect.isgeneratorfunction(g))\n",
    "True"
);
crate::runtime_case!(
    generator_getgeneratorstate,
    "import inspect\ndef g():\n yield 1\nit = g()\nnext(it)\nprint(inspect.getgeneratorstate(it))\n",
    "GEN_SUSPENDED"
);
crate::runtime_case!(
    generator_throw_while_running,
    "def g():\n yield 1\n yield 2\nit = g()\nnext(it)\ntry:\n it.throw(RuntimeError('e'))\nexcept RuntimeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    generator_close_raises,
    "def g():\n try:\n  yield 1\n except GeneratorExit:\n  print('exit')\n  raise\nit = g()\nnext(it)\nit.close()\n",
    "exit"
);
crate::runtime_case!(
    generator_yield_in_finally,
    "def g():\n try:\n  yield 1\n finally:\n  pass\nprint(list(g()))\n",
    "[1]"
);
crate::runtime_case!(
    generator_nested_yield_from,
    "def a():\n yield 1\ndef b():\n yield from a()\n yield 2\ndef c():\n yield from b()\nprint(list(c()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    generator_send_after_close,
    "def g():\n yield 1\nit = g()\nnext(it)\nit.close()\ntry:\n it.send(1)\n print('ok')\nexcept StopIteration:\n print('stop')\n",
    "stop"
);
crate::runtime_case!(
    generator_iter_id,
    "def g():\n yield 1\nprint(iter(g()) is not None)\n",
    "True"
);
crate::runtime_case!(
    generator_reentrant_iter,
    "def g():\n yield 1\nprint(list(g()) == list(g()))\n",
    "True"
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
crate::runtime_case!(
    generator_return_in_try,
    "def g():\n try:\n  yield 1\n  return\n except:\n  pass\n yield 2\nprint(list(g()))\n",
    "[1]"
);
crate::runtime_case!(
    generator_yield_from_empty,
    "def g():\n yield from []\n yield 1\nprint(list(g()))\n",
    "[1]"
);
crate::runtime_case!(
    generator_throw_stopiteration_value,
    "def g():\n if False:\n  yield 1\n return 42\nit = g()\ntry:\n next(it)\nexcept StopIteration as e:\n print(e.value)\n",
    "42"
);
crate::runtime_case!(
    generator_pep380_yield_from_expr,
    "def subgen():\n yield 2\ndef g():\n x = yield from subgen()\n yield x\nprint(list(g()))\n",
    "[2, None]"
);
crate::runtime_case!(
    generator_class_based,
    "class G:\n def __iter__(self):\n  return self\n def __next__(self):\n  raise StopIteration\nprint(list(G()) == [])\n",
    "True"
);
crate::runtime_case!(
    generator_iterable_not_iterator,
    "def g():\n yield 1\nprint(hasattr(g(), '__iter__'))\n",
    "True"
);
crate::runtime_case!(
    generator_next_method,
    "def g():\n yield 1\nprint(hasattr(g(), '__next__'))\n",
    "True"
);
crate::runtime_case!(
    generator_send_not_started,
    "def g():\n x = yield 1\n yield x\ntry:\n g().send(1)\n print('ok')\nexcept TypeError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    generator_multi_yield_from,
    "def g():\n yield from [1]\n yield from [2]\nprint(list(g()))\n",
    "[1, 2]"
);
crate::runtime_case!(
    generator_exception_in_yield_from,
    "def inner():\n yield 1\n raise ValueError\ndef outer():\n yield from inner()\ntry:\n list(outer())\nexcept ValueError:\n print('err')\n",
    "err"
);
crate::runtime_case!(
    generator_gi_frame,
    "def g():\n yield 1\nit = g()\nprint(it.gi_frame is not None)\n",
    "True"
);
crate::runtime_case!(
    generator_gi_running,
    "def g():\n yield 1\nit = g()\nprint(it.gi_running)\n",
    "False"
);
crate::runtime_case!(
    generator_gi_yieldfrom,
    "def inner():\n yield 1\ndef outer():\n yield from inner()\nit = outer()\nnext(it)\nprint(it.gi_yieldfrom is not None)\n",
    "True"
);
crate::runtime_case!(
    generator_gi_code,
    "def g():\n yield 1\nprint(g().__code__.co_name)\n",
    "g"
);

crate::compile_case!(generator_async, "async def ag():\n yield 1\n");
crate::compile_case!(
    generator_yield_from_await,
    "async def ag():\n yield from async_iter()\n"
);
crate::compile_case!(generator_pep479, "def g():\n return\n yield 1\n");
crate::compile_case!(
    generator_throw_generator,
    "def g():\n yield 1\nit = g()\nit.throw(GeneratorExit)\n"
);
crate::compile_case!(generator_asend, "async def ag():\n yield 1\n");
