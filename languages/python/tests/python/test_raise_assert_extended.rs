//! assert, raise, and debugging statement patterns.

crate::runtime_case!(assert_true, "assert True\nprint('ok')\n", "ok");
crate::runtime_case!(assert_expression, "assert 2 + 2 == 4\nprint('ok')\n", "ok");
crate::runtime_case!(
    assert_with_message,
    "x = 5\nassert x > 0, 'positive'\nprint(x)\n",
    "5"
);
crate::runtime_case!(
    assert_fails_caught,
    "try:\n assert False\nexcept AssertionError:\n print('fail')\n",
    "fail"
);
crate::runtime_case!(
    assert_message_preserved,
    "try:\n assert 0, 'zero'\nexcept AssertionError as e:\n print(str(e))\n",
    "zero"
);
crate::runtime_case!(
    raise_valueerror,
    "try:\n raise ValueError('bad')\nexcept ValueError as e:\n print(e.args[0])\n",
    "bad"
);
crate::runtime_case!(
    raise_typeerror,
    "try:\n raise TypeError\nexcept TypeError:\n print('te')\n",
    "te"
);
crate::runtime_case!(
    raise_runtimeerror,
    "try:\n raise RuntimeError('rt')\nexcept RuntimeError:\n print('rt')\n",
    "rt"
);
crate::runtime_case!(
    raise_keyerror,
    "try:\n raise KeyError('k')\nexcept KeyError as e:\n print(str(e))\n",
    "'k'"
);
crate::runtime_case!(
    raise_indexerror,
    "try:\n raise IndexError\nexcept IndexError:\n print('ie')\n",
    "ie"
);
crate::runtime_case!(
    raise_stopiteration,
    "try:\n raise StopIteration\nexcept StopIteration:\n print('si')\n",
    "si"
);
crate::runtime_case!(
    raise_notimplemented,
    "try:\n raise NotImplementedError\nexcept NotImplementedError:\n print('ni')\n",
    "ni"
);
crate::runtime_case!(
    raise_assertion_via_assert,
    "try:\n assert 1 == 2\nexcept AssertionError:\n print('assert')\n",
    "assert"
);
crate::runtime_case!(
    raise_from_chain,
    "try:\n try:\n  int('x')\n except ValueError as e:\n  raise RuntimeError('wrap') from e\nexcept RuntimeError:\n print('wrap')\n",
    "wrap"
);
crate::runtime_case!(
    raise_cause_attr,
    "try:\n try:\n  raise ValueError('v')\n except ValueError as e:\n  raise RuntimeError('r') from e\nexcept RuntimeError as e:\n print(type(e.__cause__).__name__)\n",
    "ValueError"
);
crate::runtime_case!(
    raise_no_cause,
    "try:\n raise ValueError()\nexcept ValueError as e:\n print(e.__cause__ is None)\n",
    "True"
);
crate::runtime_case!(
    raise_args_tuple,
    "try:\n raise ValueError(1, 2, 3)\nexcept ValueError as e:\n print(len(e.args))\n",
    "3"
);
crate::runtime_case!(
    raise_custom_class,
    "class E(Exception):\n pass\ntry:\n raise E('custom')\nexcept E as e:\n print(str(e))\n",
    "custom"
);
crate::runtime_case!(
    raise_in_function,
    "def f():\n raise ValueError('fn')\ntry:\n f()\nexcept ValueError as e:\n print(e.args[0])\n",
    "fn"
);
crate::runtime_case!(
    raise_reraise_bare,
    "try:\n try:\n  raise ValueError('inner')\n except ValueError:\n  raise\nexcept ValueError as e:\n print(e.args[0])\n",
    "inner"
);
crate::runtime_case!(
    assert_in_function,
    "def f():\n assert 1 == 1\n return 'ok'\nprint(f())\n",
    "ok"
);
crate::runtime_case!(
    assert_in_loop,
    "for i in range(2):\n assert i >= 0\nprint('done')\n",
    "done"
);
crate::runtime_case!(
    raise_systemexit,
    "try:\n raise SystemExit(42)\nexcept SystemExit as e:\n print(e.code)\n",
    "42"
);
crate::runtime_case!(
    raise_keyboardinterrupt,
    "try:\n raise KeyboardInterrupt()\nexcept KeyboardInterrupt:\n print('ki')\n",
    "ki"
);
crate::runtime_case!(
    raise_generator_exit,
    "try:\n raise GeneratorExit()\nexcept GeneratorExit:\n print('ge')\n",
    "ge"
);
crate::runtime_case!(
    raise_overflow,
    "try:\n raise OverflowError()\nexcept OverflowError:\n print('of')\n",
    "of"
);
crate::runtime_case!(
    raise_zero_division,
    "try:\n 1/0\nexcept ZeroDivisionError:\n print('zd')\n",
    "zd"
);
crate::runtime_case!(
    raise_attribute,
    "try:\n None.x\nexcept AttributeError:\n print('ae')\n",
    "ae"
);
crate::runtime_case!(
    raise_name_error,
    "try:\n undefined_xyz\nexcept NameError:\n print('ne')\n",
    "ne"
);
crate::runtime_case!(
    raise_import_error,
    "try:\n raise ImportError('mod')\nexcept ImportError as e:\n print(e.args[0])\n",
    "mod"
);
crate::runtime_case!(
    raise_os_error,
    "try:\n raise OSError('os')\nexcept OSError:\n print('os')\n",
    "os"
);
crate::runtime_case!(
    raise_unicode_error,
    "try:\n b'\\xff'.decode('ascii')\nexcept UnicodeError:\n print('ue')\n",
    "ue"
);
crate::runtime_case!(assert_bool_context, "assert [1]\nprint('ok')\n", "ok");
crate::runtime_case!(
    assert_empty_fails,
    "try:\n assert []\nexcept AssertionError:\n print('empty')\n",
    "empty"
);
crate::runtime_case!(
    raise_exception_base,
    "try:\n raise Exception('base')\nexcept Exception as e:\n print(e.args[0])\n",
    "base"
);
crate::runtime_case!(
    raise_lookup_error,
    "try:\n raise LookupError()\nexcept LookupError:\n print('le')\n",
    "le"
);
crate::runtime_case!(
    raise_arithmetic_error,
    "try:\n raise ArithmeticError()\nexcept ArithmeticError:\n print('ae')\n",
    "ae"
);
crate::runtime_case!(
    raise_warning_not_exception,
    "import warnings\nwarnings.warn('w')\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    raise_in_class_init,
    "class C:\n def __init__(self):\n  raise ValueError('init')\ntry:\n C()\nexcept ValueError:\n print('init')\n",
    "init"
);
crate::runtime_case!(
    raise_in_property,
    "class C:\n @property\n def x(self):\n  raise ValueError('prop')\ntry:\n C().x\nexcept ValueError:\n print('prop')\n",
    "prop"
);
crate::runtime_case!(
    assert_comparison_chain,
    "a, b, c = 1, 2, 3\nassert a < b < c\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(
    raise_with_none_arg,
    "try:\n raise ValueError(None)\nexcept ValueError as e:\n print(e.args[0] is None)\n",
    "True"
);
crate::runtime_case!(
    raise_stopiteration_value,
    "try:\n raise StopIteration(99)\nexcept StopIteration as e:\n print(e.value)\n",
    "99"
);
crate::runtime_case!(
    raise_multiple_types_catch,
    "try:\n raise KeyError()\nexcept (KeyError, ValueError):\n print('caught')\n",
    "caught"
);
crate::runtime_case!(
    raise_unbound_local_after,
    "try:\n raise ValueError()\nexcept ValueError:\n msg = 'handled'\nprint(msg)\n",
    "handled"
);
crate::runtime_case!(
    assert_identity,
    "a = [1]\nassert a is a\nprint('ok')\n",
    "ok"
);
crate::runtime_case!(assert_membership, "assert 1 in [1, 2]\nprint('ok')\n", "ok");
crate::runtime_case!(
    raise_in_finally,
    "try:\n try:\n  raise ValueError('a')\n finally:\n  raise RuntimeError('b')\nexcept RuntimeError as e:\n print(e.args[0])\n",
    "b"
);
crate::runtime_case!(
    assert_docstring_preserved,
    "def f():\n  '''doc'''\n assert True\nprint(f.__doc__)\n",
    "doc"
);

crate::compile_case!(
    raise_from_none,
    "try:\n raise ValueError() from None\nexcept ValueError:\n pass\n"
);
crate::compile_case!(assert_debug, "assert __debug__\n");
crate::compile_case!(
    raise_chained_context,
    "try:\n 1/0\nexcept ZeroDivisionError as e:\n raise ValueError() from e\n"
);
crate::compile_case!(
    raise_in_generator,
    "def g():\n yield 1\n raise ValueError()\nlist(g())\n"
);
