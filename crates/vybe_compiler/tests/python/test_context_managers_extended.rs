//! Context managers: with, nested, else, custom __enter__/__exit__, contextlib.

use crate::helpers::*;

crate::runtime_case!(
    with_basic,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nwith CM() as c:\n print('in')\n",
    "in"
);
crate::runtime_case!(
    with_return_binding,
    "class CM:\n def __enter__(self):\n  return 42\n def __exit__(self, *a):\n  pass\nwith CM() as x:\n print(x)\n",
    "42"
);
crate::runtime_case!(
    with_exit_called,
    "log = []\nclass CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  log.append('exit')\nwith CM():\n pass\nprint(log)\n",
    "['exit']"
);
crate::runtime_case!(
    with_suppress_exception,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc_type, exc, tb):\n  return True\nwith CM():\n raise ValueError()\nprint('after')\n",
    "after"
);
crate::runtime_case!(
    with_propagate_exception,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return False\ntry:\n with CM():\n  raise ValueError('e')\nexcept ValueError:\n print('caught')\n",
    "caught"
);
crate::runtime_case!(
    with_nested,
    "class CM:\n def __init__(self, v):\n  self.v = v\n def __enter__(self):\n  return self.v\n def __exit__(self, *a):\n  pass\nwith CM(1) as a, CM(2) as b:\n print(a, b)\n",
    "1 2"
);
crate::runtime_case!(
    with_multiple_sequential,
    "class CM:\n def __init__(self, v):\n  self.v = v\n def __enter__(self):\n  return self.v\n def __exit__(self, *a):\n  pass\nwith CM(1) as a:\n with CM(2) as b:\n  print(a + b)\n",
    "3"
);
crate::runtime_case!(
    with_file_like,
    "class F:\n def __enter__(self):\n  self.buf = []\n  return self\n def write(self, s):\n  self.buf.append(s)\n def __exit__(self, *a):\n  pass\nwith F() as f:\n f.write('x')\nprint(f.buf)\n",
    "['x']"
);
crate::runtime_case!(
    contextlib_suppress,
    "from contextlib import suppress\nwith suppress(ValueError):\n raise ValueError()\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    contextlib_nullcontext,
    "from contextlib import nullcontext\nwith nullcontext() as x:\n print(x is None)\n",
    "True"
);
crate::runtime_case!(
    contextlib_closing,
    "from contextlib import closing\nclass R:\n def close(self):\n  self.closed = True\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nr = R()\nwith closing(r) as x:\n pass\nprint(r.closed)\n",
    "True"
);
crate::runtime_case!(
    generator_context_manager,
    "from contextlib import contextmanager\n@contextmanager\ndef cm():\n yield 7\nwith cm() as v:\n print(v)\n",
    "7"
);
crate::runtime_case!(
    generator_context_finally,
    "log = []\nfrom contextlib import contextmanager\n@contextmanager\ndef cm():\n try:\n  yield 1\n finally:\n  log.append('fin')\nwith cm():\n pass\nprint(log)\n",
    "['fin']"
);
crate::runtime_case!(
    with_enter_raises,
    "class CM:\n def __enter__(self):\n  raise RuntimeError('enter')\n def __exit__(self, *a):\n  pass\ntry:\n with CM():\n  pass\nexcept RuntimeError:\n print('enter_err')\n",
    "enter_err"
);
crate::runtime_case!(
    with_body_raises_exit_runs,
    "log = []\nclass CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  log.append('exit')\n  return False\ntry:\n with CM():\n  raise TypeError()\nexcept TypeError:\n print(log)\n",
    "['exit']"
);
crate::runtime_case!(
    with_return_inside,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\ndef f():\n with CM():\n  return 9\nprint(f())\n",
    "9"
);
crate::runtime_case!(
    with_break_inside,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nfor i in range(2):\n with CM():\n  if i:\n   break\nprint('done')\n",
    "done"
);
crate::runtime_case!(
    with_continue_inside,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nfor i in range(2):\n with CM():\n  if i == 0:\n   continue\n  print(i)\n",
    "1"
);
crate::runtime_case!(
    with_as_tuple_unpack,
    "class CM:\n def __enter__(self):\n  return (1, 2)\n def __exit__(self, *a):\n  pass\nwith CM() as (a, b):\n print(a + b)\n",
    "3"
);
crate::runtime_case!(
    with_reentrant_same,
    "class CM:\n def __enter__(self):\n  self.n = getattr(self, 'n', 0) + 1\n  return self.n\n def __exit__(self, *a):\n  pass\nc = CM()\nwith c as a:\n with c as b:\n  print(a, b)\n",
    "1 2"
);
crate::runtime_case!(
    with_exit_receives_exc_info,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc_type, exc, tb):\n  print(exc_type is not None)\n  return True\nwith CM():\n raise ValueError()\n",
    "True"
);
crate::runtime_case!(
    with_exit_none_exc,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc_type, exc, tb):\n  print(exc_type is None)\nwith CM():\n pass\n",
    "True"
);
crate::runtime_case!(
    contextlib_redirect_stdout,
    "from contextlib import redirect_stdout\nimport io\nbuf = io.StringIO()\nwith redirect_stdout(buf):\n print('hidden')\nprint(len(buf.getvalue()) > 0)\n",
    "True"
);
crate::runtime_case!(
    contextlib_exitstack,
    "from contextlib import ExitStack\nclass CM:\n def __enter__(self):\n  return 1\n def __exit__(self, *a):\n  pass\nwith ExitStack() as stack:\n v = stack.enter_context(CM())\n print(v)\n",
    "1"
);
crate::runtime_case!(
    with_magic_methods_bool,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return None\nwith CM():\n print('yes')\n",
    "yes"
);
crate::runtime_case!(
    with_enter_return_self_mutate,
    "class CM:\n def __enter__(self):\n  self.v = 5\n  return self\n def __exit__(self, *a):\n  pass\nwith CM() as c:\n print(c.v)\n",
    "5"
);
crate::runtime_case!(
    with_open_mock,
    "class F:\n def __enter__(self):\n  return ['line']\n def __exit__(self, *a):\n  pass\nwith F() as lines:\n print(lines[0])\n",
    "line"
);
crate::runtime_case!(
    with_try_finally_order,
    "log = []\nclass CM:\n def __enter__(self):\n  log.append('enter')\n  return self\n def __exit__(self, *a):\n  log.append('exit')\ntry:\n with CM():\n  log.append('body')\nfinally:\n  log.append('fin')\nprint(log)\n",
    "['enter', 'body', 'exit', 'fin']"
);
crate::runtime_case!(
    with_else_clause,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return False\ntry:\n with CM():\n  pass\nelse:\n print('else')\nexcept:\n pass\n",
    "else"
);
crate::runtime_case!(
    with_exception_suppressed_no_else,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return True\nmatched = False\ntry:\n with CM():\n  raise ValueError()\nelse:\n matched = True\nprint(matched)\n",
    "False"
);
crate::runtime_case!(
    contextmanager_yield_once,
    "from contextlib import contextmanager\n@contextmanager\ndef cm():\n yield 'x'\nwith cm() as v:\n print(v)\n",
    "x"
);
crate::runtime_case!(
    contextmanager_catch_yield,
    "from contextlib import contextmanager\n@contextmanager\ndef cm():\n try:\n  yield 1\n except ValueError:\n  print('caught')\nwith cm():\n raise ValueError()\n",
    "caught"
);
crate::runtime_case!(
    with_del_resource,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\n def __del__(self):\n  pass\nwith CM():\n print(1)\n",
    "1"
);
crate::runtime_case!(
    with_lambda_enter,
    "class CM:\n def __enter__(self):\n  return (lambda: 9)()\n def __exit__(self, *a):\n  pass\nwith CM() as v:\n print(v)\n",
    "9"
);
crate::runtime_case!(
    with_class_decorator,
    "def deco(cls):\n return cls\n@deco\nclass CM:\n def __enter__(self):\n  return 1\n def __exit__(self, *a):\n  pass\nwith CM() as v:\n print(v)\n",
    "1"
);
crate::runtime_case!(
    with_thread_lock_like,
    "class L:\n def __init__(self):\n  self.locked = False\n def __enter__(self):\n  self.locked = True\n  return self\n def __exit__(self, *a):\n  self.locked = False\nl = L()\nwith l:\n print(l.locked)\n",
    "True"
);
crate::runtime_case!(
    with_timing_pattern,
    "class Timer:\n def __enter__(self):\n  self.t0 = 0\n  return self\n def __exit__(self, *a):\n  pass\nwith Timer() as t:\n print(hasattr(t, 't0'))\n",
    "True"
);
crate::runtime_case!(
    with_suppress_keyerror,
    "from contextlib import suppress\nwith suppress(KeyError):\n {}['x']\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    with_chained_managers,
    "class CM:\n def __init__(self, v):\n  self.v = v\n def __enter__(self):\n  return self.v\n def __exit__(self, *a):\n  pass\nwith CM(1) as a, CM(2) as b, CM(3) as c:\n print(a + b + c)\n",
    "6"
);
crate::runtime_case!(
    with_enter_attribute_error,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nwith CM() as c:\n print(hasattr(c, '__enter__'))\n",
    "True"
);
crate::runtime_case!(
    with_generator_close,
    "from contextlib import contextmanager\n@contextmanager\ndef cm():\n try:\n  yield 1\n finally:\n  print('closed')\nwith cm():\n pass\n",
    "closed"
);
crate::runtime_case!(
    with_return_from_finally,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\ndef f():\n try:\n  with CM():\n   return 1\n finally:\n  pass\nprint(f())\n",
    "1"
);
crate::runtime_case!(
    with_nested_suppress,
    "from contextlib import suppress\nwith suppress(ValueError):\n with suppress(TypeError):\n  raise ValueError()\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    with_custom_exception_in_exit,
    "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc_type, exc, tb):\n  if exc_type:\n   print('saw')\n  return True\nwith CM():\n raise OSError()\n",
    "saw"
);
crate::runtime_case!(
    with_enter_none,
    "class CM:\n def __enter__(self):\n  return None\n def __exit__(self, *a):\n  pass\nwith CM() as x:\n print(x is None)\n",
    "True"
);
crate::runtime_case!(
    with_long_body,
    "class CM:\n def __enter__(self):\n  return 0\n def __exit__(self, *a):\n  pass\nwith CM() as s:\n for i in range(3):\n  s += i\n print(s)\n",
    "3"
);

crate::compile_case!(with_async_context, "class CM:\n async def __aenter__(self):\n  return self\n async def __aexit__(self, *a):\n  pass\n");
crate::compile_case!(contextlib_asynccontextmanager, "from contextlib import asynccontextmanager\n@asynccontextmanager\nasync def cm():\n yield 1\n");
crate::compile_case!(with_open_builtin, "with open('/dev/null', 'w') as f:\n pass\n");
crate::compile_case!(contextlib_chdir, "from contextlib import chdir\n");
crate::compile_case!(contextlib_suppress_multiple, "from contextlib import suppress\nwith suppress(ValueError, TypeError):\n pass\n");
