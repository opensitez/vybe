use crate::helpers::{run_python_one, run_python};

#[test]
fn try_except_catches_zero_division() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept ZeroDivisionError:\n print('caught')\n"),
        "caught"
    );
}

#[test]
fn try_except_catches_type_error() {
    assert_eq!(
        run_python_one("try:\n 'a' + 1\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn try_except_catches_value_error() {
    assert_eq!(
        run_python_one("try:\n int('x')\nexcept ValueError:\n print('value')\n"),
        "value"
    );
}

#[test]
fn try_except_catches_index_error() {
    assert_eq!(
        run_python_one("try:\n [1][5]\nexcept IndexError:\n print('index')\n"),
        "index"
    );
}

#[test]
fn try_except_catches_key_error() {
    assert_eq!(
        run_python_one("try:\n {}['k']\nexcept KeyError:\n print('key')\n"),
        "key"
    );
}

#[test]
fn try_except_else_runs_when_no_error() {
    assert_eq!(
        run_python_one("try:\n x = 1\nexcept:\n print('no')\nelse:\n print('else')\n"),
        "else"
    );
}

#[test]
fn try_except_finally_always_runs() {
    assert_eq!(
        run_python_one("try:\n print('try')\nexcept:\n pass\nfinally:\n print('fin')\n"),
        "try\nfin"
    );
}

#[test]
fn try_except_finally_on_exception() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept ZeroDivisionError:\n print('ex')\nfinally:\n print('fin')\n"),
        "ex\nfin"
    );
}

#[test]
fn try_except_bare_except_catches_any() {
    assert_eq!(
        run_python_one("try:\n raise Exception('x')\nexcept:\n print('any')\n"),
        "any"
    );
}

#[test]
fn try_except_exception_base_type() {
    assert_eq!(
        run_python_one("try:\n raise ValueError('bad')\nexcept Exception:\n print('exc')\n"),
        "exc"
    );
}

#[test]
fn try_except_as_binds_message() {
    assert_eq!(
        run_python_one("try:\n raise ValueError('oops')\nexcept ValueError as e:\n print(str(e))\n"),
        "oops"
    );
}

#[test]
fn try_except_multiple_handlers_first_match() {
    assert_eq!(
        run_python_one(
            "try:\n raise TypeError('t')\nexcept ValueError:\n print('v')\nexcept TypeError:\n print('t')\n"
        ),
        "t"
    );
}

#[test]
fn try_except_multiple_handlers_second_match() {
    assert_eq!(
        run_python_one(
            "try:\n raise ValueError('v')\nexcept TypeError:\n print('t')\nexcept ValueError:\n print('v')\n"
        ),
        "v"
    );
}

#[test]
fn try_except_nested_inner_caught() {
    assert_eq!(
        run_python_one(
            "try:\n try:\n  1/0\n except ZeroDivisionError:\n  print('inner')\nexcept:\n print('outer')\n"
        ),
        "inner"
    );
}

#[test]
fn try_except_nested_outer_catches() {
    assert_eq!(
        run_python_one(
            "try:\n try:\n  raise KeyError\n except ValueError:\n  print('inner')\nexcept KeyError:\n print('outer')\n"
        ),
        "outer"
    );
}

#[test]
fn try_except_reraise_not_caught_by_wrong_type() {
    let lines = run_python(
        "try:\n try:\n  raise ValueError('x')\n except TypeError:\n  print('wrong')\nexcept ValueError:\n print('right')\n"
    );
    assert_eq!(lines, vec!["right"]);
}

#[test]
fn try_except_else_skipped_on_exception() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept ZeroDivisionError:\n print('ex')\nelse:\n print('else')\n"),
        "ex"
    );
}

#[test]
fn try_except_return_in_try() {
    assert_eq!(
        run_python_one("def f():\n try:\n  return 1\n except:\n  return 2\nprint(f())\n"),
        "1"
    );
}

#[test]
fn try_except_return_in_except() {
    assert_eq!(
        run_python_one("def f():\n try:\n  1/0\n except:\n  return 2\nprint(f())\n"),
        "2"
    );
}

#[test]
fn try_except_break_in_try() {
    assert_eq!(
        run_python_one("for i in range(3):\n try:\n  if i == 1:\n   break\n except:\n  pass\nprint(i)\n"),
        "1"
    );
}

#[test]
fn try_except_continue_in_except() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(3):\n try:\n  if i == 1:\n   raise ValueError\n  out.append(i)\n except ValueError:\n  continue\nprint(out)\n"
        ),
        "[0, 2]"
    );
}

#[test]
fn try_except_attribute_error() {
    assert_eq!(
        run_python_one("try:\n None.x\nexcept AttributeError:\n print('attr')\n"),
        "attr"
    );
}

#[test]
fn try_except_name_error() {
    assert_eq!(
        run_python_one("try:\n no_such_name\nexcept NameError:\n print('name')\n"),
        "name"
    );
}

#[test]
fn try_except_runtime_error() {
    assert_eq!(
        run_python_one("try:\n raise RuntimeError('rt')\nexcept RuntimeError:\n print('rt')\n"),
        "rt"
    );
}

#[test]
fn try_except_stop_iteration() {
    assert_eq!(
        run_python_one("try:\n next(iter([]))\nexcept StopIteration:\n print('stop')\n"),
        "stop"
    );
}

#[test]
fn try_except_assertion_error() {
    assert_eq!(
        run_python_one("try:\n assert False\nexcept AssertionError:\n print('assert')\n"),
        "assert"
    );
}

#[test]
fn try_except_lookup_error_parent() {
    assert_eq!(
        run_python_one("try:\n [][0]\nexcept LookupError:\n print('lookup')\n"),
        "lookup"
    );
}

#[test]
fn try_except_arithmetic_error_parent() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept ArithmeticError:\n print('arith')\n"),
        "arith"
    );
}

#[test]
fn try_except_os_error_subclass_simulated() {
    assert_eq!(
        run_python_one("try:\n raise OSError('disk')\nexcept OSError as e:\n print('os')\n"),
        "os"
    );
}

#[test]
fn try_except_finally_overrides_no_return_path() {
    assert_eq!(
        run_python_one("def f():\n try:\n  return 1\n finally:\n  print('f')\nprint(f())\n"),
        "f\n1"
    );
}

#[test]
fn try_except_assign_in_try_used_after() {
    assert_eq!(
        run_python_one("try:\n x = 5\nexcept:\n x = 0\nprint(x)\n"),
        "5"
    );
}

#[test]
fn try_except_assign_in_except_used_after() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept:\n x = 9\nprint(x)\n"),
        "9"
    );
}

#[test]
fn try_except_tuple_unpack_in_except() {
    assert_eq!(
        run_python_one("try:\n a, b = [1]\nexcept ValueError:\n print('unpack')\n"),
        "unpack"
    );
}

#[test]
fn try_except_with_else_assign() {
    assert_eq!(
        run_python_one("try:\n n = 2\nexcept:\n n = 0\nelse:\n print(n * 3)\n"),
        "6"
    );
}

#[test]
fn try_except_raise_without_args_reraises() {
    assert_eq!(
        run_python_one(
            "try:\n try:\n  raise ValueError('z')\n except ValueError:\n  raise\nexcept ValueError as e:\n print(str(e))\n"
        ),
        "z"
    );
}

#[test]
fn try_except_catch_after_successful_try_block() {
    assert_eq!(
        run_python_one("try:\n result = 2 + 2\nexcept:\n result = 0\nprint(result)\n"),
        "4"
    );
}

#[test]
fn try_except_function_call_in_try() {
    assert_eq!(
        run_python_one("def boom():\n raise ValueError('fn')\ntry:\n boom()\nexcept ValueError:\n print('fn')\n"),
        "fn"
    );
}

#[test]
fn try_except_loop_with_periodic_errors() {
    assert_eq!(
        run_python_one(
            "count = 0\nfor i in range(3):\n try:\n  if i == 1:\n   raise ValueError\n  count += 1\n except ValueError:\n  pass\nprint(count)\n"
        ),
        "2"
    );
}

#[test]
fn try_except_base_exception_not_caught_by_exception() {
    assert_eq!(
        run_python_one("try:\n raise KeyboardInterrupt\nexcept Exception:\n print('exc')\nexcept BaseException:\n print('base')\n"),
        "base"
    );
}

#[test]
fn try_except_specific_before_general() {
    assert_eq!(
        run_python_one("try:\n int('x')\nexcept ValueError:\n print('specific')\nexcept Exception:\n print('general')\n"),
        "specific"
    );
}

#[test]
fn try_except_empty_try_success() {
    assert_eq!(
        run_python_one("try:\n pass\nexcept:\n print('no')\nelse:\n print('ok')\n"),
        "ok"
    );
}

#[test]
fn try_except_finally_with_return_in_try() {
    assert_eq!(
        run_python_one("def f():\n try:\n  return 'r'\n finally:\n  pass\nprint(f())\n"),
        "r"
    );
}

#[test]
fn try_except_len_on_wrong_type() {
    assert_eq!(
        run_python_one("try:\n len(42)\nexcept TypeError:\n print('len')\n"),
        "len"
    );
}

#[test]
fn try_except_iter_on_int() {
    assert_eq!(
        run_python_one("try:\n iter(123)\nexcept TypeError:\n print('iter')\n"),
        "iter"
    );
}

#[test]
fn try_except_chained_handling_preserves_flow() {
    assert_eq!(
        run_python_one(
            "def work():\n try:\n  return 1/0\n except ZeroDivisionError:\n  return 'fixed'\nprint(work())\n"
        ),
        "fixed"
    );
}
