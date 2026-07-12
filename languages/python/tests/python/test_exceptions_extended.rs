//! Exception hierarchy, raise from, chained exceptions, else/finally interaction.

crate::runtime_case!(
    except_valueerror_type,
    "try:\n int('x')\nexcept ValueError:\n print('ve')\n",
    "ve"
);
crate::runtime_case!(
    except_type_error,
    "try:\n {} + 1\nexcept TypeError:\n print('te')\n",
    "te"
);
crate::runtime_case!(
    except_key_error,
    "try:\n {}['k']\nexcept KeyError:\n print('ke')\n",
    "ke"
);
crate::runtime_case!(
    except_index_error,
    "try:\n [][0]\nexcept IndexError:\n print('ie')\n",
    "ie"
);
crate::runtime_case!(
    except_attribute_error,
    "try:\n None.x\nexcept AttributeError:\n print('ae')\n",
    "ae"
);
crate::runtime_case!(
    except_zero_division,
    "try:\n 1 / 0\nexcept ZeroDivisionError:\n print('zd')\n",
    "zd"
);
crate::runtime_case!(
    except_stop_iteration,
    "try:\n next(iter([]))\nexcept StopIteration:\n print('si')\n",
    "si"
);
crate::runtime_case!(
    except_base_exception,
    "try:\n raise KeyboardInterrupt()\nexcept BaseException:\n print('be')\n",
    "be"
);
crate::runtime_case!(
    except_tuple_types,
    "try:\n int('a')\nexcept (ValueError, TypeError):\n print('tuple')\n",
    "tuple"
);
crate::runtime_case!(
    except_as_binding,
    "try:\n raise ValueError('msg')\nexcept ValueError as e:\n print(str(e))\n",
    "msg"
);
crate::runtime_case!(
    except_else_runs,
    "try:\n x = 1\nexcept:\n x = 2\nelse:\n print(x)\n",
    "1"
);
crate::runtime_case!(
    except_else_skipped_on_error,
    "try:\n raise ValueError()\nexcept ValueError:\n print('caught')\nelse:\n print('else')\n",
    "caught"
);
crate::runtime_case!(
    finally_always_runs,
    "try:\n print('try')\nfinally:\n print('fin')\n",
    "try\nfin"
);
crate::runtime_case!(
    finally_on_exception,
    "try:\n raise ValueError()\nexcept ValueError:\n print('ex')\nfinally:\n print('fin')\n",
    "ex\nfin"
);
crate::runtime_case!(
    raise_reraise,
    "try:\n try:\n  raise ValueError('inner')\n except ValueError:\n  raise\nexcept ValueError as e:\n print('outer')\n",
    "outer"
);
crate::runtime_case!(
    raise_from_cause,
    "try:\n try:\n  int('x')\n except ValueError as e:\n  raise RuntimeError('wrap') from e\nexcept RuntimeError:\n print('wrapped')\n",
    "wrapped"
);
crate::runtime_case!(
    raise_bare_inside_except,
    "try:\n raise TypeError('t')\nexcept TypeError:\n try:\n  raise ValueError('v')\n except ValueError:\n  print('v')\n",
    "v"
);
crate::runtime_case!(
    except_hierarchy_lookup,
    "try:\n raise LookupError()\nexcept KeyError:\n print('ke')\nexcept LookupError:\n print('le')\n",
    "le"
);
crate::runtime_case!(
    except_os_error_subclass,
    "try:\n raise OSError('os')\nexcept OSError as e:\n print('os')\n",
    "os"
);
crate::runtime_case!(
    except_assertion_error,
    "try:\n assert False\nexcept AssertionError:\n print('assert')\n",
    "assert"
);
crate::runtime_case!(
    except_not_implemented,
    "try:\n raise NotImplementedError()\nexcept NotImplementedError:\n print('ni')\n",
    "ni"
);
crate::runtime_case!(
    except_overflow,
    "try:\n raise OverflowError()\nexcept OverflowError:\n print('of')\n",
    "of"
);
crate::runtime_case!(
    except_unicode_error,
    "try:\n b'\\xff'.decode('ascii')\nexcept UnicodeError:\n print('ue')\n",
    "ue"
);
crate::runtime_case!(
    except_generator_exit,
    "try:\n raise GeneratorExit()\nexcept GeneratorExit:\n print('ge')\n",
    "ge"
);
crate::runtime_case!(
    except_system_exit_code,
    "try:\n raise SystemExit(3)\nexcept SystemExit as e:\n print(e.code)\n",
    "3"
);
crate::runtime_case!(
    except_exception_group,
    "try:\n raise Exception('plain')\nexcept Exception:\n print('ex')\n",
    "ex"
);
crate::runtime_case!(
    try_nested_inner,
    "try:\n try:\n  1/0\n except ZeroDivisionError:\n  print('inner')\nexcept:\n print('outer')\n",
    "inner"
);
crate::runtime_case!(
    try_nested_outer_catches,
    "try:\n try:\n  raise KeyError()\n except TypeError:\n  print('no')\nexcept KeyError:\n print('outer')\n",
    "outer"
);
crate::runtime_case!(
    finally_return_override,
    "def f():\n try:\n  return 1\n finally:\n  pass\nprint(f())\n",
    "1"
);
crate::runtime_case!(
    except_name_error,
    "try:\n print(undefined_name_xyz)\nexcept NameError:\n print('ne')\n",
    "ne"
);
crate::runtime_case!(
    raise_value_with_args,
    "try:\n raise ValueError(1, 2)\nexcept ValueError as e:\n print(len(e.args))\n",
    "2"
);
crate::runtime_case!(
    exception_str_repr,
    "e = ValueError('bad')\nprint('ValueError' in type(e).__name__)\n",
    "True"
);
crate::runtime_case!(
    exception_isinstance,
    "print(isinstance(ValueError(), Exception))\n",
    "True"
);
crate::runtime_case!(
    exception_issubclass,
    "print(issubclass(KeyError, LookupError))\n",
    "True"
);
crate::runtime_case!(
    except_break_propagation,
    "for i in range(1):\n try:\n  break\n except:\n  pass\nprint('done')\n",
    "done"
);
crate::runtime_case!(
    except_continue_propagation,
    "for i in range(2):\n try:\n  if i == 0:\n   continue\n  print(i)\n except:\n  pass\n",
    "1"
);
crate::runtime_case!(
    finally_break_suppresses,
    "for i in range(1):\n try:\n  break\n finally:\n  pass\nprint('after')\n",
    "after"
);
crate::runtime_case!(
    except_runtime_error,
    "try:\n raise RuntimeError('rt')\nexcept RuntimeError as e:\n print(e.args[0])\n",
    "rt"
);
crate::runtime_case!(
    raise_stop_iteration_ctor,
    "try:\n raise StopIteration(42)\nexcept StopIteration as e:\n print(e.value)\n",
    "42"
);
crate::runtime_case!(
    except_warning_not_caught_by_exception,
    "import warnings\nprint('warn_ok')\n",
    "warn_ok"
);
crate::runtime_case!(
    except_else_finally_order,
    "try:\n pass\nexcept:\n pass\nelse:\n print('e')\nfinally:\n print('f')\n",
    "e\nf"
);
crate::runtime_case!(
    multiple_except_first_match,
    "try:\n [][0]\nexcept (KeyError, IndexError):\n print('matched')\n",
    "matched"
);
crate::runtime_case!(
    except_unbound_local_after,
    "x = 1\ntry:\n raise ValueError()\nexcept ValueError:\n y = 2\nprint(y)\n",
    "2"
);

crate::compile_case!(
    except_bare_except,
    "try:\n raise ValueError()\nexcept:\n pass\n"
);
crate::compile_case!(raise_not_implemented, "raise NotImplementedError('todo')\n");
crate::compile_case!(
    except_exception_chaining_context,
    "try:\n 1/0\nexcept ZeroDivisionError as e:\n raise ValueError() from e\n"
);
crate::compile_case!(
    try_finally_return,
    "def f():\n try:\n  return 1\n finally:\n  return 2\n"
);
crate::compile_case!(
    except_match_case,
    "try:\n raise ValueError()\nexcept ValueError:\n match 1:\n  case 1:\n   pass\n"
);
