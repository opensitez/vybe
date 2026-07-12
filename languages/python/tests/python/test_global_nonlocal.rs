use crate::helpers::run_python_one;

#[test]
fn global_read_without_assignment() {
    assert_eq!(
        run_python_one("g = 1\ndef f():\n return g\nprint(f())\n"),
        "1"
    );
}

#[test]
fn global_declaration_writes_outer() {
    assert_eq!(
        run_python_one("g = 1\ndef f():\n global g\n g = 2\nf()\nprint(g)\n"),
        "2"
    );
}

#[test]
fn global_without_decl_creates_local() {
    assert_eq!(
        run_python_one("g = 1\ndef f():\n g = 9\n return g\nprint(f(), g)\n"),
        "9 1"
    );
}

#[test]
fn nonlocal_updates_enclosing() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 1\n def inner():\n  nonlocal x\n  x = 2\n inner()\n return x\nprint(outer())\n"
        ),
        "2"
    );
}

#[test]
fn nonlocal_without_global() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 0\n def inner():\n  nonlocal x\n  x += 1\n inner()\n return x\nprint(outer())\n"
        ),
        "1"
    );
}

#[test]
fn nested_nonlocal_two_levels() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 1\n def mid():\n  def inner():\n   nonlocal x\n   x = 3\n  inner()\n mid()\n return x\nprint(outer())\n"
        ),
        "3"
    );
}

#[test]
fn global_and_nonlocal_different_names() {
    assert_eq!(
        run_python_one(
            "g = 0\ndef outer():\n x = 1\n def inner():\n  global g\n  nonlocal x\n  g = 5\n  x = 6\n inner()\n return x\nprint(outer(), g)\n"
        ),
        "6 5"
    );
}

#[test]
fn closure_read_enclosing_no_nonlocal() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 10\n def inner():\n  return x\n return inner()\nprint(outer())\n"
        ),
        "10"
    );
}

#[test]
fn closure_assign_without_nonlocal_local_shadow() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 1\n def inner():\n  x = 2\n  return x\n return inner(), x\nprint(outer())\n"
        ),
        "(2, 1)"
    );
}

#[test]
fn global_list_mutate_no_decl_needed() {
    assert_eq!(
        run_python_one("items = [1]\ndef f():\n items.append(2)\nf()\nprint(items)\n"),
        "[1, 2]"
    );
}

#[test]
fn global_rebind_list_needs_global() {
    assert_eq!(
        run_python_one("items = [1]\ndef f():\n global items\n items = [9]\nf()\nprint(items)\n"),
        "[9]"
    );
}

#[test]
fn nonlocal_rebind_list() {
    assert_eq!(
        run_python_one(
            "def outer():\n items = [1]\n def inner():\n  nonlocal items\n  items = [2]\n inner()\n return items\nprint(outer())\n"
        ),
        "[2]"
    );
}

#[test]
fn global_in_nested_function() {
    assert_eq!(
        run_python_one(
            "count = 0\ndef outer():\n def inner():\n  global count\n  count += 1\n inner()\nouter()\nprint(count)\n"
        ),
        "1"
    );
}

#[test]
fn nonlocal_in_loop_closure() {
    assert_eq!(
        run_python_one(
            "def outer():\n total = 0\n def add(n):\n  nonlocal total\n  total += n\n for i in range(3):\n  add(i)\n return total\nprint(outer())\n"
        ),
        "3"
    );
}

#[test]
fn global_module_level_function() {
    assert_eq!(
        run_python_one("def set_x():\n global x\n x = 7\nset_x()\nprint(x)\n"),
        "7"
    );
}

#[test]
fn nonlocal_multiple_names() {
    assert_eq!(
        run_python_one(
            "def outer():\n a, b = 1, 2\n def inner():\n  nonlocal a, b\n  a, b = 3, 4\n inner()\n return a, b\nprint(outer())\n"
        ),
        "(3, 4)"
    );
}

#[test]
fn global_read_before_write_same_function() {
    assert_eq!(
        run_python_one("g = 5\ndef f():\n global g\n return g + 1\nprint(f())\n"),
        "6"
    );
}

#[test]
fn nonlocal_read_before_write() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 0\n def inner():\n  nonlocal x\n  x = x + 1\n  return x\n return inner()\nprint(outer())\n"
        ),
        "1"
    );
}

#[test]
fn closure_capture_default_arg() {
    assert_eq!(
        run_python_one(
            "def outer(x=10):\n def inner():\n  return x\n return inner()\nprint(outer())\n"
        ),
        "10"
    );
}

#[test]
fn closure_capture_loop_variable_without_nonlocal_bug_pattern() {
    assert_eq!(
        run_python_one(
            "funcs = []\nfor i in range(2):\n funcs.append(lambda: i)\nprint(funcs[0](), funcs[1]())\n"
        ),
        "1 1"
    );
}

#[test]
fn closure_capture_loop_with_default_fix() {
    assert_eq!(
        run_python_one(
            "funcs = [lambda x=i: x for i in range(2)]\nprint(funcs[0](), funcs[1]())\n"
        ),
        "0 1"
    );
}

#[test]
fn global_builtin_not_shadowed() {
    assert_eq!(
        run_python_one("def f():\n return len([1, 2])\nprint(f())\n"),
        "2"
    );
}

#[test]
fn local_shadows_global_read() {
    assert_eq!(
        run_python_one("x = 1\ndef f():\n x = 2\n return x\nprint(f())\n"),
        "2"
    );
}

#[test]
fn global_explicit_same_value() {
    assert_eq!(
        run_python_one("x = 3\ndef f():\n global x\n x = 3\nprint(f(), x)\n"),
        "None 3"
    );
}

#[test]
fn nonlocal_not_visible_at_module() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 1\n def inner():\n  nonlocal x\n  x = 2\n inner()\n return x\nprint(outer())\n"
        ),
        "2"
    );
}

#[test]
fn nested_functions_separate_locals() {
    assert_eq!(
        run_python_one(
            "def a():\n x = 1\n def b():\n  x = 9\n  return x\n return b(), x\nprint(a())\n"
        ),
        "(9, 1)"
    );
}

#[test]
fn global_dict_update_nested() {
    assert_eq!(
        run_python_one(
            "state = {'n': 0}\ndef inc():\n state['n'] += 1\ninc()\nprint(state['n'])\n"
        ),
        "1"
    );
}

#[test]
fn nonlocal_counter_factory() {
    assert_eq!(
        run_python_one(
            "def make_counter():\n n = 0\n def inc():\n  nonlocal n\n  n += 1\n  return n\n return inc\nc = make_counter()\nprint(c(), c())\n"
        ),
        "1 2"
    );
}

#[test]
fn global_reassign_int_chain() {
    assert_eq!(
        run_python_one(
            "n = 0\ndef a():\n global n\n n = 1\ndef b():\n global n\n n = 2\na()\nb()\nprint(n)\n"
        ),
        "2"
    );
}

#[test]
fn nonlocal_in_try_block() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 0\n try:\n  def inner():\n   nonlocal x\n   x = 5\n  inner()\n except:\n  pass\n return x\nprint(outer())\n"
        ),
        "5"
    );
}

#[test]
fn global_in_try_block() {
    assert_eq!(
        run_python_one(
            "g = 0\ntry:\n def f():\n  global g\n  g = 8\n f()\nexcept:\n pass\nprint(g)\n"
        ),
        "8"
    );
}

#[test]
fn closure_with_nonlocal_and_global_mix() {
    assert_eq!(
        run_python_one(
            "g = 0\ndef outer():\n x = 1\n def inner():\n  global g\n  nonlocal x\n  g = 10\n  x = 11\n  return x\n return inner()\nprint(outer(), g)\n"
        ),
        "11 10"
    );
}

#[test]
fn local_variable_not_leaking() {
    assert_eq!(
        run_python_one("def f():\n local = 99\n return local\nprint(f())\n"),
        "99"
    );
}

#[test]
fn global_name_matches_local_param_no_conflict() {
    assert_eq!(
        run_python_one("g = 1\ndef f(g):\n return g\nprint(f(5))\n"),
        "5"
    );
}

#[test]
fn nonlocal_only_affects_nearest_enclosing() {
    assert_eq!(
        run_python_one(
            "def a():\n x = 1\n def b():\n  x = 2\n  def c():\n   nonlocal x\n   x = 3\n  c()\n  return x\n return b()\nprint(a())\n"
        ),
        "3"
    );
}

#[test]
fn global_function_reference() {
    assert_eq!(
        run_python_one(
            "def real():\n return 1\ndef caller():\n global real\n return real()\nprint(caller())\n"
        ),
        "1"
    );
}

#[test]
fn nonlocal_del_not_allowed_use_reassign() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 1\n def inner():\n  nonlocal x\n  x = None\n inner()\n return x\nprint(outer())\n"
        ),
        "None"
    );
}

#[test]
fn module_level_implicit_global() {
    assert_eq!(
        run_python_one("counter = 0\ncounter += 1\nprint(counter)\n"),
        "1"
    );
}

#[test]
fn nested_global_same_name_inner() {
    assert_eq!(
        run_python_one(
            "x = 'outer'\ndef outer():\n def inner():\n  global x\n  x = 'inner'\n inner()\nouter()\nprint(x)\n"
        ),
        "inner"
    );
}

#[test]
fn closure_returns_function_with_nonlocal() {
    assert_eq!(
        run_python_one(
            "def outer():\n n = 0\n def inc():\n  nonlocal n\n  n += 1\n  return n\n return inc\nf = outer()\nprint(f(), f())\n"
        ),
        "1 2"
    );
}

#[test]
fn global_list_rebind_vs_mutate() {
    assert_eq!(
        run_python_one(
            "a = [1]\ndef mutate():\n a.append(2)\ndef rebind():\n global a\n a = [9]\nmutate()\nprint(a)\nrebind()\nprint(a)\n"
        ),
        "[1, 2]\n[9]"
    );
}

#[test]
fn nonlocal_shared_via_wrapper_function() {
    assert_eq!(
        run_python_one(
            "def outer():\n x = 1\n def get():\n  return x\n def set(v):\n  nonlocal x\n  x = v\n set(5)\n return get()\nprint(outer())\n"
        ),
        "5"
    );
}
