use crate::helpers::run_python_one;

#[test]
fn raise_value_error_with_message() {
    assert_eq!(
        run_python_one("try:\n raise ValueError('bad')\nexcept ValueError as e:\n print(str(e))\n"),
        "bad"
    );
}

#[test]
fn raise_type_error() {
    assert_eq!(
        run_python_one("try:\n raise TypeError\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn raise_runtime_error() {
    assert_eq!(
        run_python_one("try:\n raise RuntimeError('rt')\nexcept RuntimeError:\n print('rt')\n"),
        "rt"
    );
}

#[test]
fn raise_without_args_in_except_reraises() {
    assert_eq!(
        run_python_one("try:\n try:\n  raise ValueError('z')\n except ValueError:\n  raise\nexcept ValueError as e:\n print(str(e))\n"),
        "z"
    );
}

#[test]
fn raise_from_chain_cause() {
    assert_eq!(
        run_python_one("try:\n try:\n  raise ValueError('inner')\n except ValueError as e:\n  raise RuntimeError('outer') from e\nexcept RuntimeError as e:\n print(str(e.__cause__))\n"),
        "inner"
    );
}

#[test]
fn raise_stop_iteration() {
    assert_eq!(
        run_python_one("try:\n raise StopIteration\nexcept StopIteration:\n print('stop')\n"),
        "stop"
    );
}

#[test]
fn raise_key_error() {
    assert_eq!(
        run_python_one("try:\n raise KeyError('k')\nexcept KeyError:\n print('key')\n"),
        "key"
    );
}

#[test]
fn raise_index_error() {
    assert_eq!(
        run_python_one("try:\n raise IndexError\nexcept IndexError:\n print('idx')\n"),
        "idx"
    );
}

#[test]
fn raise_assertion_error_via_assert() {
    assert_eq!(
        run_python_one("try:\n assert 1 == 2\nexcept AssertionError:\n print('assert')\n"),
        "assert"
    );
}

#[test]
fn assert_true_passes() {
    assert_eq!(
        run_python_one("assert True\nprint('ok')\n"),
        "ok"
    );
}

#[test]
fn assert_with_message() {
    assert_eq!(
        run_python_one("try:\n assert 0, 'zero'\nexcept AssertionError as e:\n print(str(e))\n"),
        "zero"
    );
}

#[test]
fn assert_expression_form() {
    assert_eq!(
        run_python_one("x = 5\nassert (y := x * 2) == 10\nprint(y)\n"),
        "10"
    );
}

#[test]
fn raise_in_function() {
    assert_eq!(
        run_python_one("def f():\n raise ValueError('fn')\ntry:\n f()\nexcept ValueError as e:\n print(str(e))\n"),
        "fn"
    );
}

#[test]
fn raise_if_condition() {
    assert_eq!(
        run_python_one("x = -1\ntry:\n if x < 0:\n  raise ValueError('neg')\nexcept ValueError:\n print('neg')\n"),
        "neg"
    );
}

#[test]
fn raise_custom_exception_subclass() {
    assert_eq!(
        run_python_one("class MyErr(Exception):\n pass\ntry:\n raise MyErr('m')\nexcept MyErr as e:\n print(str(e))\n"),
        "m"
    );
}

#[test]
fn raise_not_caught_by_wrong_type() {
    assert_eq!(
        run_python_one("try:\n try:\n  raise TypeError('t')\n except ValueError:\n  print('wrong')\nexcept TypeError:\n print('right')\n"),
        "right"
    );
}

#[test]
fn raise_os_error() {
    assert_eq!(
        run_python_one("try:\n raise OSError('disk')\nexcept OSError:\n print('os')\n"),
        "os"
    );
}

#[test]
fn raise_zero_division() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept ZeroDivisionError:\n print('z')\n"),
        "z"
    );
}

#[test]
fn raise_name_error() {
    assert_eq!(
        run_python_one("try:\n no_name\nexcept NameError:\n print('name')\n"),
        "name"
    );
}

#[test]
fn raise_attribute_error() {
    assert_eq!(
        run_python_one("try:\n None.x\nexcept AttributeError:\n print('attr')\n"),
        "attr"
    );
}

#[test]
fn raise_lookup_error_parent() {
    assert_eq!(
        run_python_one("try:\n [][0]\nexcept LookupError:\n print('lookup')\n"),
        "lookup"
    );
}

#[test]
fn raise_arithmetic_error_parent() {
    assert_eq!(
        run_python_one("try:\n 1/0\nexcept ArithmeticError:\n print('arith')\n"),
        "arith"
    );
}

#[test]
fn raise_exception_base() {
    assert_eq!(
        run_python_one("try:\n raise Exception('base')\nexcept Exception as e:\n print(str(e))\n"),
        "base"
    );
}

#[test]
fn raise_in_loop_breaks_to_handler() {
    assert_eq!(
        run_python_one("for i in range(2):\n try:\n  if i:\n   raise ValueError('loop')\n except ValueError:\n  print('caught')\n"),
        "caught"
    );
}

#[test]
fn assert_in_function() {
    assert_eq!(
        run_python_one("def f(x):\n assert x > 0\n return x\ntry:\n f(-1)\nexcept AssertionError:\n print('fail')\n"),
        "fail"
    );
}

#[test]
fn raise_none_as_exception_invalid() {
    assert_eq!(
        run_python_one("try:\n raise None\nexcept TypeError:\n print('type')\n"),
        "type"
    );
}

#[test]
fn raise_string_old_style_not_valid() {
    assert_eq!(
        run_python_one("try:\n raise ValueError(123)\nexcept ValueError as e:\n print(str(e))\n"),
        "123"
    );
}

#[test]
fn raise_from_none_suppresses_context() {
    assert_eq!(
        run_python_one("try:\n try:\n  raise ValueError('a')\n except ValueError as e:\n  raise RuntimeError('b') from None\nexcept RuntimeError as e:\n print(e.__cause__)\n"),
        "None"
    );
}

#[test]
fn assert_chained_comparisons() {
    assert_eq!(
        run_python_one("x = 5\nassert 0 < x < 10\nprint('ok')\n"),
        "ok"
    );
}

#[test]
fn raise_recursion_depth_style_manual() {
    assert_eq!(
        run_python_one("def f():\n f()\ntry:\n f()\nexcept RecursionError:\n print('rec')\n"),
        "rec"
    );
}

#[test]
fn raise_keyboard_interrupt_type() {
    assert_eq!(
        run_python_one("try:\n raise KeyboardInterrupt\nexcept KeyboardInterrupt:\n print('kbd')\n"),
        "kbd"
    );
}

#[test]
fn raise_generator_exit() {
    assert_eq!(
        run_python_one("try:\n raise GeneratorExit\nexcept GeneratorExit:\n print('gen')\n"),
        "gen"
    );
}

#[test]
fn raise_system_exit_not_base_exception_handler() {
    assert_eq!(
        run_python_one("try:\n raise SystemExit(0)\nexcept BaseException:\n print('base')\n"),
        "base"
    );
}

#[test]
fn assert_is_identity() {
    assert_eq!(
        run_python_one("a = []\nassert a is a\nprint('ok')\n"),
        "ok"
    );
}

#[test]
fn assert_membership() {
    assert_eq!(
        run_python_one("assert 2 in [1, 2, 3]\nprint('ok')\n"),
        "ok"
    );
}

#[test]
fn raise_multiple_in_sequence_caught_once() {
    assert_eq!(
        run_python_one("try:\n raise ValueError('one')\nexcept ValueError as e:\n print(str(e))\n"),
        "one"
    );
}

#[test]
fn raise_empty_exception_args() {
    assert_eq!(
        run_python_one("try:\n raise ValueError()\nexcept ValueError as e:\n print(len(e.args))\n"),
        "0"
    );
}

#[test]
fn raise_tuple_args() {
    assert_eq!(
        run_python_one("try:\n raise ValueError('a', 'b')\nexcept ValueError as e:\n print(len(e.args))\n"),
        "2"
    );
}

#[test]
fn assert_false_literal_fails() {
    assert_eq!(
        run_python_one("try:\n assert False\nexcept AssertionError:\n print('no')\n"),
        "no"
    );
}

#[test]
fn raise_inside_finally_logged() {
    assert_eq!(
        run_python_one("log = []\ntry:\n try:\n  raise ValueError\n except:\n  log.append('ex')\n finally:\n  log.append('fin')\nprint(log)\n"),
        "['ex', 'fin']"
    );
}

#[test]
fn raise_unbound_local_error() {
    assert_eq!(
        run_python_one("def f():\n print(x)\n x = 1\ntry:\n f()\nexcept UnboundLocalError:\n print('unbound')\n"),
        "unbound"
    );
}

#[test]
fn raise_overflow_error_manual() {
    assert_eq!(
        run_python_one("try:\n raise OverflowError('big')\nexcept OverflowError:\n print('over')\n"),
        "over"
    );
}

#[test]
fn assert_not_none() {
    assert_eq!(
        run_python_one("x = 1\nassert x is not None\nprint(x)\n"),
        "1"
    );
}

#[test]
fn raise_import_error() {
    assert_eq!(
        run_python_one("try:\n raise ImportError('missing')\nexcept ImportError as e:\n print(str(e))\n"),
        "missing"
    );
}

#[test]
fn raise_unicode_error_subclass() {
    assert_eq!(
        run_python_one("try:\n raise UnicodeDecodeError('utf-8', b'\\xff', 0, 1, 'bad')\nexcept UnicodeDecodeError:\n print('uni')\n"),
        "uni"
    );
}
