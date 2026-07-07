use crate::helpers::run_python_one;

#[test]
fn finally_runs_after_try_success() {
    assert_eq!(
        run_python_one("out = []\ntry:\n out.append(1)\nfinally:\n out.append(2)\nprint(out)\n"),
        "[1, 2]"
    );
}

#[test]
fn finally_runs_after_except() {
    assert_eq!(
        run_python_one(
            "out = []\ntry:\n 1 / 0\nexcept ZeroDivisionError:\n out.append('e')\nfinally:\n out.append('f')\nprint(out)\n"
        ),
        "['e', 'f']"
    );
}

#[test]
fn else_runs_only_without_exception() {
    assert_eq!(
        run_python_one(
            "out = []\ntry:\n out.append(1)\nexcept:\n out.append('x')\nelse:\n out.append('else')\nprint(out)\n"
        ),
        "[1, 'else']"
    );
}

#[test]
fn else_skipped_when_exception_handled() {
    assert_eq!(
        run_python_one(
            "out = []\ntry:\n raise ValueError('x')\nexcept ValueError:\n out.append('caught')\nelse:\n out.append('else')\nprint(out)\n"
        ),
        "['caught']"
    );
}

#[test]
fn finally_runs_before_else() {
    assert_eq!(
        run_python_one(
            "out = []\ntry:\n out.append('t')\nexcept:\n pass\nelse:\n out.append('e')\nfinally:\n out.append('f')\nprint(out)\n"
        ),
        "['t', 'e', 'f']"
    );
}

#[test]
fn try_except_specific_type_only() {
    assert_eq!(
        run_python_one("try:\n int('x')\nexcept ValueError:\n print('ve')\n"),
        "ve"
    );
}

#[test]
fn try_except_tuple_of_types() {
    assert_eq!(
        run_python_one("try:\n int('x')\nexcept (ValueError, TypeError):\n print('ok')\n"),
        "ok"
    );
}

#[test]
fn raise_reraise_after_except() {
    assert_eq!(
        run_python_one(
            "try:\n try:\n  raise ValueError('inner')\n except ValueError:\n  raise RuntimeError('outer') from None\nexcept RuntimeError as e:\n print(type(e).__name__)\n"
        ),
        "RuntimeError"
    );
}

#[test]
fn except_as_binds_exception() {
    assert_eq!(
        run_python_one("try:\n raise ValueError('msg')\nexcept ValueError as e:\n print(str(e))\n"),
        "msg"
    );
}

#[test]
fn finally_return_suppressed_by_exception_in_try() {
    assert_eq!(
        run_python_one("def f():\n try:\n  return 1\n finally:\n  return 2\nprint(f())\n"),
        "2"
    );
}

#[test]
fn try_nested_inner_except() {
    assert_eq!(
        run_python_one(
            "try:\n try:\n  1/0\n except ZeroDivisionError:\n  print('inner')\nexcept:\n print('outer')\n"
        ),
        "inner"
    );
}

#[test]
fn finally_in_nested_try() {
    assert_eq!(
        run_python_one(
            "out = []\ntry:\n try:\n  out.append(1)\n finally:\n  out.append(2)\nfinally:\n out.append(3)\nprint(out)\n"
        ),
        "[1, 2, 3]"
    );
}

#[test]
fn except_catches_base_exception_subclass() {
    assert_eq!(
        run_python_one("try:\n raise KeyboardInterrupt()\nexcept BaseException:\n print('base')\n"),
        "base"
    );
}

#[test]
fn try_with_break_in_loop() {
    assert_eq!(
        run_python_one(
            "for _ in range(3):\n try:\n  print('x')\n  break\n finally:\n  print('y')\n"
        ),
        "x\ny"
    );
}

#[test]
fn try_with_continue_in_loop() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(3):\n try:\n  if i == 1:\n   continue\n  out.append(i)\n finally:\n  pass\nprint(out)\n"
        ),
        "[0, 2]"
    );
}

#[test]
fn assert_with_message() {
    assert_eq!(
        run_python_one(
            "try:\n assert 1 == 2, 'bad'\nexcept AssertionError as e:\n print('AssertionError')\n"
        ),
        "AssertionError"
    );
}

#[test]
fn assert_true_passes_silently() {
    assert_eq!(run_python_one("assert True\nprint('ok')\n"), "ok");
}

#[test]
fn except_finally_else_order_on_success() {
    assert_eq!(
        run_python_one(
            "log = []\ntry:\n log.append('t')\nexcept:\n log.append('e')\nelse:\n log.append('el')\nfinally:\n log.append('f')\nprint(log)\n"
        ),
        "['t', 'el', 'f']"
    );
}

#[test]
fn try_finally_variable_assignment_visible() {
    assert_eq!(
        run_python_one("x = 0\ntry:\n x = 1\nfinally:\n x += 10\nprint(x)\n"),
        "11"
    );
}

#[test]
fn except_multiple_handlers_first_match() {
    assert_eq!(
        run_python_one(
            "try:\n raise TypeError('t')\nexcept ValueError:\n print('v')\nexcept TypeError:\n print('t')\n"
        ),
        "t"
    );
}

#[test]
fn raise_value_error_no_arg() {
    assert_eq!(
        run_python_one("try:\n raise ValueError\nexcept ValueError:\n print('ok')\n"),
        "ok"
    );
}

#[test]
fn try_return_in_try_finally_still_runs() {
    assert_eq!(
        run_python_one("def f():\n try:\n  return 'a'\n finally:\n  pass\nprint(f())\n"),
        "a"
    );
}

#[test]
fn except_pass_swallows() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept ZeroDivisionError:\n pass\nprint('after')\n"),
        "after"
    );
}

#[test]
fn finally_overwrites_return_value() {
    assert_eq!(
        run_python_one("def g():\n try:\n  return 1\n finally:\n  return 9\nprint(g())\n"),
        "9"
    );
}

#[test]
fn try_except_in_function_propagates_uncaught() {
    assert_eq!(
        run_python_one(
            "def h():\n raise RuntimeError('x')\ntry:\n h()\nexcept RuntimeError:\n print('caught')\n"
        ),
        "caught"
    );
}

#[test]
fn else_not_run_when_break_in_try() {
    assert_eq!(
        run_python_one(
            "for _ in range(1):\n try:\n  break\n else:\n  print('else')\nprint('done')\n"
        ),
        "done"
    );
}

#[test]
fn finally_closes_resource_pattern() {
    assert_eq!(
        run_python_one(
            "closed = []\ntry:\n closed.append('open')\nfinally:\n closed.append('close')\nprint(closed)\n"
        ),
        "['open', 'close']"
    );
}

#[test]
fn except_as_exception_args() {
    assert_eq!(
        run_python_one(
            "try:\n raise ValueError(1, 2, 3)\nexcept ValueError as e:\n print(len(e.args))\n"
        ),
        "3"
    );
}

#[test]
fn try_nested_finally_only_inner_on_inner_break() {
    assert_eq!(
        run_python_one(
            "out = []\nfor _ in range(1):\n try:\n  try:\n   break\n  finally:\n   out.append('i')\n finally:\n  out.append('o')\nprint(out)\n"
        ),
        "['i', 'o']"
    );
}

#[test]
fn raise_from_preserves_context_type() {
    assert_eq!(
        run_python_one(
            "try:\n try:\n  1/0\n except ZeroDivisionError as e:\n  raise ValueError('wrap') from e\nexcept ValueError:\n print('ValueError')\n"
        ),
        "ValueError"
    );
}

#[test]
fn try_else_finally_with_return_in_else() {
    assert_eq!(
        run_python_one(
            "def f():\n try:\n  pass\n else:\n  return 5\n finally:\n  pass\nprint(f())\n"
        ),
        "5"
    );
}

#[test]
fn except_handles_key_error() {
    assert_eq!(
        run_python_one("try:\n {}['x']\nexcept KeyError:\n print('key')\n"),
        "key"
    );
}

#[test]
fn except_handles_index_error() {
    assert_eq!(
        run_python_one("try:\n [1][9]\nexcept IndexError:\n print('index')\n"),
        "index"
    );
}

#[test]
fn except_handles_type_error() {
    assert_eq!(
        run_python_one("try:\n 'a' + 1\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn finally_runs_on_continue_in_loop() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(2):\n try:\n  if i == 0:\n   continue\n  out.append(i)\n finally:\n  out.append(9)\nprint(out)\n"
        ),
        "[9, 1, 9]"
    );
}

#[test]
fn try_except_else_finally_empty_try() {
    assert_eq!(
        run_python_one(
            "log = []\ntry:\n pass\nexcept:\n log.append('e')\nelse:\n log.append('ok')\nfinally:\n log.append('f')\nprint(log)\n"
        ),
        "['ok', 'f']"
    );
}

#[test]
fn nested_except_bare_except_catches_all() {
    assert_eq!(
        run_python_one(
            "try:\n try:\n  raise ValueError\n except:\n  print('inner')\nexcept:\n print('outer')\n"
        ),
        "inner"
    );
}

#[test]
fn try_finally_with_break_in_finally_not_allowed_use_pattern() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(2):\n try:\n  out.append(i)\n finally:\n  if i == 1:\n   pass\nprint(out)\n"
        ),
        "[0, 1]"
    );
}

#[test]
fn except_exception_name_bound_in_local_scope() {
    assert_eq!(
        run_python_one(
            "def f():\n try:\n  raise ValueError('z')\n except ValueError as err:\n  return str(err)\nprint(f())\n"
        ),
        "z"
    );
}

#[test]
fn try_return_finally_mutates_outer() {
    assert_eq!(
        run_python_one(
            "box = {'v': 0}\ndef f():\n try:\n  return 1\n finally:\n  box['v'] = 9\nf()\nprint(box['v'])\n"
        ),
        "9"
    );
}
