//! async/await, async for/with, asyncio basics — runtime where possible.

crate::runtime_case!(
    async_def_callable,
    "async def f():\n return 1\nprint(callable(f))\n",
    "True"
);
crate::runtime_case!(
    coroutine_object_type_name,
    "async def f():\n return 1\nc = f()\nprint(type(c).__name__)\n",
    "coroutine"
);
crate::runtime_case!(
    async_function_name_preserved,
    "async def hello():\n pass\nprint(hello.__name__)\n",
    "hello"
);
crate::runtime_case!(
    await_in_async_def_compile_check,
    "async def f():\n return 1\nprint('sync_ok')\n",
    "sync_ok"
);
crate::runtime_case!(
    async_lambda_not_valid_use_async_def,
    "async def f():\n x = 1\n return x\nprint('defined')\n",
    "defined"
);
crate::runtime_case!(
    async_method_in_class,
    "class C:\n async def m(self):\n  return 1\nprint(callable(C.m))\n",
    "True"
);
crate::runtime_case!(
    async_gen_function,
    "async def ag():\n yield 1\nprint(callable(ag))\n",
    "True"
);
crate::runtime_case!(
    async_gen_object_name,
    "async def ag():\n yield 1\ng = ag()\nprint(type(g).__name__)\n",
    "async_generator"
);
crate::runtime_case!(
    asyncio_module_import,
    "import asyncio\nprint(hasattr(asyncio, 'run'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_get_event_loop_policy,
    "import asyncio\nprint(hasattr(asyncio, 'get_event_loop_policy'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_iscoroutine,
    "import asyncio\nasync def f():\n pass\nprint(asyncio.iscoroutine(f()))\n",
    "True"
);
crate::runtime_case!(
    asyncio_iscoroutinefunction,
    "import asyncio\nasync def f():\n pass\ndef g():\n pass\nprint(asyncio.iscoroutinefunction(f))\n",
    "True"
);
crate::runtime_case!(
    asyncio_iscoroutinefunction_false,
    "import asyncio\ndef g():\n pass\nprint(asyncio.iscoroutinefunction(g))\n",
    "False"
);
crate::runtime_case!(
    asyncio_future_class,
    "import asyncio\nprint(hasattr(asyncio, 'Future'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_task_class,
    "import asyncio\nprint(hasattr(asyncio, 'Task'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_lock_class,
    "import asyncio\nprint(hasattr(asyncio, 'Lock'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_event_class,
    "import asyncio\nprint(hasattr(asyncio, 'Event'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_queue_class,
    "import asyncio\nprint(hasattr(asyncio, 'Queue'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_sleep_exists,
    "import asyncio\nprint(callable(asyncio.sleep))\n",
    "True"
);
crate::runtime_case!(
    asyncio_gather_exists,
    "import asyncio\nprint(callable(asyncio.gather))\n",
    "True"
);
crate::runtime_case!(
    asyncio_create_task_exists,
    "import asyncio\nprint(callable(asyncio.create_task))\n",
    "True"
);
crate::runtime_case!(
    asyncio_wait_exists,
    "import asyncio\nprint(callable(asyncio.wait))\n",
    "True"
);
crate::runtime_case!(
    asyncio_shield_exists,
    "import asyncio\nprint(callable(asyncio.shield))\n",
    "True"
);
crate::runtime_case!(
    asyncio_timeout_exists,
    "import asyncio\nprint(hasattr(asyncio, 'timeout'))\n",
    "True"
);
crate::runtime_case!(
    async_with_statement_compile,
    "class CM:\n async def __aenter__(self):\n  return self\n async def __aexit__(self, *a):\n  pass\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    async_for_statement_compile,
    "async def ag():\n yield 1\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    await_expression_parse,
    "async def f():\n x = 1\n return x\nprint(f.__name__)\n",
    "f"
);
crate::runtime_case!(
    async_nested_function,
    "async def outer():\n async def inner():\n  return 2\n return inner\nprint(callable(outer))\n",
    "True"
);
crate::runtime_case!(
    async_default_arg,
    "async def f(x=1):\n return x\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    async_varargs,
    "async def f(*args):\n return len(args)\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    async_kwargs,
    "async def f(**kwargs):\n return kwargs\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    async_raise_in_body,
    "async def f():\n raise ValueError('e')\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    async_try_except,
    "async def f():\n try:\n  raise ValueError()\n except ValueError:\n  return 1\n return 0\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    async_return_value,
    "async def f():\n return 42\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    async_yield_from_syntax,
    "async def ag():\n yield 1\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    asyncio_coroutines_module,
    "import asyncio.coroutines\nprint(hasattr(asyncio.coroutines, 'iscoroutine'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_streams_module,
    "import asyncio\nprint(hasattr(asyncio, 'StreamReader'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_subprocess_module,
    "import asyncio\nprint(hasattr(asyncio, 'create_subprocess_exec'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_all_tasks_exists,
    "import asyncio\nprint(hasattr(asyncio, 'all_tasks'))\n",
    "True"
);
crate::runtime_case!(
    asyncio_current_task_exists,
    "import asyncio\nprint(hasattr(asyncio, 'current_task'))\n",
    "True"
);
crate::compile_case!(
    async_comprehension_syntax,
    "async def f():\n return [i async for i in aiter([])]\n"
);
crate::runtime_case!(
    asyncio_run_callable,
    "import asyncio\nprint(callable(asyncio.run))\n",
    "True"
);
crate::runtime_case!(
    asyncio_get_running_loop_raises_sync,
    "import asyncio\ntry:\n asyncio.get_running_loop()\n print('has')\nexcept RuntimeError:\n print('none')\n",
    "none"
);
crate::runtime_case!(
    async_generator_asend_exists,
    "async def ag():\n yield 1\ng = ag()\nprint(hasattr(g, 'asend'))\n",
    "True"
);
crate::runtime_case!(
    async_generator_athrow_exists,
    "async def ag():\n yield 1\ng = ag()\nprint(hasattr(g, 'athrow'))\n",
    "True"
);
crate::runtime_case!(
    async_generator_aclose_exists,
    "async def ag():\n yield 1\ng = ag()\nprint(hasattr(g, 'aclose'))\n",
    "True"
);
crate::runtime_case!(
    coroutine_close_exists,
    "async def f():\n pass\nc = f()\nprint(hasattr(c, 'close'))\n",
    "True"
);
crate::runtime_case!(
    coroutine_throw_exists,
    "async def f():\n pass\nc = f()\nprint(hasattr(c, 'throw'))\n",
    "True"
);

crate::compile_case!(asyncio_run_simple, "import asyncio\nasync def main():\n return 1\nasyncio.run(main())\n");
crate::compile_case!(async_with_contextlib, "from contextlib import asynccontextmanager\n@asynccontextmanager\nasync def cm():\n yield 1\n");
crate::compile_case!(async_gather_two, "import asyncio\nasync def f():\n return 1\nasync def g():\n return 2\nasyncio.run(asyncio.gather(f(), g()))\n");
crate::compile_case!(async_wait_for, "import asyncio\nasync def f():\n return 1\nasyncio.run(asyncio.wait_for(f(), timeout=1))\n");
crate::compile_case!(async_task_group, "import asyncio\nasync def main():\n async with asyncio.TaskGroup() as tg:\n  pass\nasyncio.run(main())\n");
