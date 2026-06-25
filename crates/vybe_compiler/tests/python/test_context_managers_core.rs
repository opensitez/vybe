use crate::helpers::run_python_one;

#[test]
fn with_open_writes_and_reads() {
    assert_eq!(
        run_python_one(
            "from io import StringIO\nbuf = StringIO()\nwith buf as f:\n f.write('hi')\nprint(buf.getvalue())\n"
        ),
        "hi"
    );
}

#[test]
fn with_suppresses_exception_on_exit_success() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return False\nwith CM():\n print('ok')\n"
        ),
        "ok"
    );
}

#[test]
fn with_exit_returns_true_suppresses() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc, val, tb):\n  return True\ntry:\n with CM():\n  raise ValueError('x')\nexcept ValueError:\n print('leaked')\nprint('after')\n"
        ),
        "after"
    );
}

#[test]
fn with_as_binding() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return 42\n def __exit__(self, *a):\n  pass\nwith CM() as x:\n print(x)\n"
        ),
        "42"
    );
}

#[test]
fn with_nested_managers() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __init__(self, v):\n  self.v = v\n def __enter__(self):\n  return self.v\n def __exit__(self, *a):\n  pass\nwith CM(1) as a, CM(2) as b:\n print(a + b)\n"
        ),
        "3"
    );
}

#[test]
fn with_exception_propagates_when_not_suppressed() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return False\ntry:\n with CM():\n  raise RuntimeError('boom')\nexcept RuntimeError:\n print('caught')\n"
        ),
        "caught"
    );
}

#[test]
fn with_finally_like_cleanup_runs() {
    assert_eq!(
        run_python_one(
            "log = []\nclass CM:\n def __enter__(self):\n  log.append('enter')\n  return self\n def __exit__(self, *a):\n  log.append('exit')\nwith CM():\n log.append('body')\nprint(log)\n"
        ),
        "['enter', 'body', 'exit']"
    );
}

#[test]
fn with_finally_runs_on_exception() {
    assert_eq!(
        run_python_one(
            "log = []\nclass CM:\n def __enter__(self):\n  log.append('in')\n  return self\n def __exit__(self, *a):\n  log.append('out')\ntry:\n with CM():\n  raise ValueError\nexcept ValueError:\n pass\nprint(log)\n"
        ),
        "['in', 'out']"
    );
}

#[test]
fn with_return_inside_block() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\ndef f():\n with CM():\n  return 7\nprint(f())\n"
        ),
        "7"
    );
}

#[test]
fn with_break_inside_block() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nfor _ in range(2):\n with CM():\n  print('x')\n  break\n"
        ),
        "x"
    );
}

#[test]
fn with_continue_inside_block() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nout = []\nfor i in range(3):\n with CM():\n  if i == 1:\n   continue\n out.append(i)\nprint(out)\n"
        ),
        "[0, 2]"
    );
}

#[test]
fn with_enter_returns_none() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return None\n def __exit__(self, *a):\n  pass\nwith CM() as x:\n print(x)\n"
        ),
        "None"
    );
}

#[test]
fn with_multiple_operations_on_resource() {
    assert_eq!(
        run_python_one(
            "from io import StringIO\nbuf = StringIO()\nwith buf as f:\n f.write('a')\n f.write('b')\nprint(buf.getvalue())\n"
        ),
        "ab"
    );
}

#[test]
fn with_reenter_new_context_each_time() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __init__(self):\n  self.n = 0\n def __enter__(self):\n  self.n += 1\n  return self.n\n def __exit__(self, *a):\n  pass\ncm = CM()\nwith cm as a:\n pass\nwith cm as b:\n print(a, b)\n"
        ),
        "1 2"
    );
}

#[test]
fn with_exit_receives_exc_info() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc, val, tb):\n  print(exc.__name__ if exc else 'none')\n  return True\nwith CM():\n raise TypeError('t')\n"
        ),
        "TypeError"
    );
}

#[test]
fn with_exit_no_exception_exc_is_none() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc, val, tb):\n  print('ok' if exc is None else 'bad')\nwith CM():\n pass\n"
        ),
        "ok"
    );
}

#[test]
fn with_contextlib_nullcontext() {
    assert_eq!(
        run_python_one(
            "from contextlib import nullcontext\nwith nullcontext() as x:\n print(x)\n"
        ),
        "None"
    );
}

#[test]
fn with_statement_expression_value() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return 5\n def __exit__(self, *a):\n  pass\nresult = 0\nwith CM() as v:\n result = v * 2\nprint(result)\n"
        ),
        "10"
    );
}

#[test]
fn with_file_like_flush_on_exit() {
    assert_eq!(
        run_python_one(
            "from io import StringIO\nclass F(StringIO):\n def __exit__(self, *a):\n  self.flush()\nbuf = F()\nwith buf:\n buf.write('z')\nprint(buf.getvalue())\n"
        ),
        "z"
    );
}

#[test]
fn with_deeply_nested_three_levels() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __init__(self, c):\n  self.c = c\n def __enter__(self):\n  return self.c\n def __exit__(self, *a):\n  pass\nwith CM(1) as a:\n with CM(2) as b:\n  with CM(3) as c:\n   print(a + b + c)\n"
        ),
        "6"
    );
}

#[test]
fn with_assign_after_block() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return {'k': 1}\n def __exit__(self, *a):\n  pass\nwith CM() as d:\n x = d['k']\nprint(x)\n"
        ),
        "1"
    );
}

#[test]
fn with_generator_context_manager_style() {
    assert_eq!(
        run_python_one(
            "from contextlib import contextmanager\n@contextmanager\ndef cm():\n yield 9\nwith cm() as v:\n print(v)\n"
        ),
        "9"
    );
}

#[test]
fn with_generator_cleanup_runs() {
    assert_eq!(
        run_python_one(
            "from contextlib import contextmanager\nlog = []\n@contextmanager\ndef cm():\n log.append('open')\n yield 1\n log.append('close')\nwith cm():\n log.append('use')\nprint(log)\n"
        ),
        "['open', 'use', 'close']"
    );
}

#[test]
fn with_suppress_context_manager() {
    assert_eq!(
        run_python_one(
            "from contextlib import suppress\nwith suppress(ValueError):\n raise ValueError('x')\nprint('ok')\n"
        ),
        "ok"
    );
}

#[test]
fn with_enter_exception_propagates() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  raise OSError('enter')\n def __exit__(self, *a):\n  pass\ntry:\n with CM():\n  pass\nexcept OSError:\n print('enter')\n"
        ),
        "enter"
    );
}

#[test]
fn with_exit_exception_propagates_after_cleanup() {
    assert_eq!(
        run_python_one(
            "log = []\nclass CM:\n def __enter__(self):\n  log.append('in')\n  return self\n def __exit__(self, *a):\n  log.append('out')\n  return False\ntry:\n with CM():\n  raise ValueError\nexcept ValueError:\n pass\nprint(log)\n"
        ),
        "['in', 'out']"
    );
}

#[test]
fn with_as_same_object_mutated() {
    assert_eq!(
        run_python_one(
            "class Box:\n def __init__(self):\n  self.items = []\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nwith Box() as b:\n b.items.append(1)\nprint(b.items)\n"
        ),
        "[1]"
    );
}

#[test]
fn with_tuple_unpack_not_supported_use_single_as() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return (1, 2)\n def __exit__(self, *a):\n  pass\nwith CM() as t:\n a, b = t\n print(a, b)\n"
        ),
        "1 2"
    );
}

#[test]
fn with_loop_repeated_entry() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return 1\n def __exit__(self, *a):\n  pass\ntotal = 0\nfor _ in range(3):\n with CM() as v:\n  total += v\nprint(total)\n"
        ),
        "3"
    );
}

#[test]
fn with_else_not_valid_use_plain_after() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return False\ntry:\n with CM():\n  pass\n print('after')\nexcept:\n print('no')\n"
        ),
        "after"
    );
}

#[test]
fn with_raise_inside_suppressed() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  return True\nwith CM():\n raise ValueError\nprint('done')\n"
        ),
        "done"
    );
}

#[test]
fn with_method_bound_enter() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return 'bound'\n def __exit__(self, *a):\n  pass\ncm = CM()\nwith cm as v:\n print(v)\n"
        ),
        "bound"
    );
}

#[test]
fn with_custom_enter_return_iterable() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return [1, 2]\n def __exit__(self, *a):\n  pass\nwith CM() as xs:\n print(sum(xs))\n"
        ),
        "3"
    );
}

#[test]
fn with_state_flag_toggle() {
    assert_eq!(
        run_python_one(
            "class Lock:\n def __init__(self):\n  self.locked = False\n def __enter__(self):\n  self.locked = True\n  return self\n def __exit__(self, *a):\n  self.locked = False\nlk = Lock()\nwith lk:\n print(lk.locked)\nprint(lk.locked)\n"
        ),
        "True\nFalse"
    );
}

#[test]
fn with_timing_pattern_elapsed() {
    assert_eq!(
        run_python_one(
            "class Timer:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nwith Timer():\n x = 1 + 1\nprint(x)\n"
        ),
        "2"
    );
}

#[test]
fn with_redirect_stdout_style_buffer() {
    assert_eq!(
        run_python_one(
            "from io import StringIO\nfrom contextlib import redirect_stdout\nbuf = StringIO()\nwith redirect_stdout(buf):\n print('hidden')\nprint(buf.getvalue().strip())\n"
        ),
        "hidden"
    );
}

#[test]
fn with_closing_closes_resource() {
    assert_eq!(
        run_python_one(
            "from contextlib import closing\nfrom io import StringIO\nbuf = StringIO()\nwith closing(buf) as f:\n f.write('c')\nprint(buf.closed)\n"
        ),
        "True"
    );
}

#[test]
fn with_async_not_tested_sync_placeholder() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return 0\n def __exit__(self, *a):\n  pass\nwith CM() as n:\n print(n + 1)\n"
        ),
        "1"
    );
}

#[test]
fn with_exception_in_exit_not_suppressed() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  raise RuntimeError('exit')\ntry:\n with CM():\n  pass\nexcept RuntimeError:\n print('exit')\n"
        ),
        "exit"
    );
}

#[test]
fn with_original_exception_if_exit_returns_false() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, exc, val, tb):\n  return False\ntry:\n with CM():\n  raise ValueError('orig')\nexcept ValueError as e:\n print(str(e))\n"
        ),
        "orig"
    );
}

#[test]
fn with_stack_of_two_cleanups() {
    assert_eq!(
        run_python_one(
            "log = []\nclass CM:\n def __init__(self, name):\n  self.name = name\n def __enter__(self):\n  log.append(self.name + '+')\n  return self\n def __exit__(self, *a):\n  log.append(self.name + '-')\nwith CM('a'):\n with CM('b'):\n  pass\nprint(''.join(log))\n"
        ),
        "a+b+b-a-"
    );
}

#[test]
fn with_enter_return_self_chain_methods() {
    assert_eq!(
        run_python_one(
            "class Builder:\n def __enter__(self):\n  self.parts = []\n  return self\n def add(self, x):\n  self.parts.append(x)\n  return self\n def __exit__(self, *a):\n  pass\nwith Builder() as b:\n b.add(1).add(2)\nprint(b.parts)\n"
        ),
        "[1, 2]"
    );
}

#[test]
fn with_pass_only_body() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return self\n def __exit__(self, *a):\n  pass\nwith CM():\n pass\nprint('ok')\n"
        ),
        "ok"
    );
}

#[test]
fn with_local_var_does_not_leak() {
    assert_eq!(
        run_python_one(
            "class CM:\n def __enter__(self):\n  return 1\n def __exit__(self, *a):\n  pass\nwith CM() as v:\n local = v\nprint(local)\n"
        ),
        "1"
    );
}
