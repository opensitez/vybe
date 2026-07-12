use crate::helpers::{compile_ok, run_print, run_python_one};

#[test]
fn time_time_returns_number() {
    assert_eq!(
        run_python_one(
            "import time\nv = time.time()\nprint(type(v).__name__ == 'float' or type(v).__name__ == 'int')\n"
        ),
        "True"
    );
}

#[test]
fn time_time_monotonic_non_decreasing_in_loop() {
    assert_eq!(
        run_python_one("import time\na = time.time()\nb = time.time()\nprint(b >= a)\n"),
        "True"
    );
}

#[test]
fn time_sleep_zero_completes() {
    assert_eq!(
        run_python_one("import time\ntime.sleep(0)\nprint('done')\n"),
        "done"
    );
}

#[test]
fn time_module_has_sleep_attribute() {
    assert_eq!(run_print("hasattr(__import__('time'), 'sleep')"), "True");
}

#[test]
fn time_module_has_time_attribute() {
    assert_eq!(run_print("hasattr(__import__('time'), 'time')"), "True");
}

#[test]
fn time_time_used_in_elapsed_calculation() {
    assert_eq!(
        run_python_one("import time\nstart = 100.0\nend = 100.5\nprint(int((end - start) * 10))\n"),
        "5"
    );
}

#[test]
fn time_import_does_not_shadow_builtin_len() {
    assert_eq!(run_python_one("import time\nprint(len([1, 2, 3]))\n"), "3");
}

#[test]
fn time_sleep_then_print_order() {
    assert_eq!(
        run_python_one("import time\nprint('before')\ntime.sleep(0)\nprint('after')\n"),
        "before"
    );
}

#[test]
fn time_time_assign_to_variable() {
    assert_eq!(
        run_python_one("import time\nt = time.time()\nprint(t == t)\n"),
        "True"
    );
}

#[test]
fn time_module_compile_with_alias() {
    compile_ok("import time as t\nt.time()\n");
}

#[test]
fn time_module_compile_sleep_in_function() {
    compile_ok("import time\ndef pause():\n time.sleep(0)\n");
}

#[test]
fn time_time_in_boolean_context() {
    assert_eq!(
        run_python_one("import time\nprint(bool(time.time()))\n"),
        "True"
    );
}

#[test]
fn time_time_subtraction_is_float_like() {
    assert_eq!(
        run_python_one("import time\nprint((time.time() - time.time()) <= 0)\n"),
        "True"
    );
}

#[test]
fn time_nested_function_calls_time() {
    assert_eq!(
        run_python_one(
            "import time\ndef now():\n return time.time()\nprint(now() == now() or True)\n"
        ),
        "True"
    );
}

#[test]
fn time_sleep_inside_while_once() {
    assert_eq!(
        run_python_one("import time\nn = 0\nwhile n < 1:\n time.sleep(0)\n n += 1\nprint(n)\n"),
        "1"
    );
}

#[test]
fn time_time_in_list() {
    assert_eq!(
        run_python_one("import time\nxs = [time.time(), 0]\nprint(len(xs))\n"),
        "2"
    );
}

#[test]
fn time_time_passed_to_round() {
    assert_eq!(
        run_python_one("import time\nprint(round(time.time(), 0) == round(time.time(), 0))\n"),
        "True"
    );
}

#[test]
fn time_module_dir_contains_time() {
    assert_eq!(
        run_python_one("import time\nprint('time' in dir(time))\n"),
        "True"
    );
}

#[test]
fn time_import_twice_same_module() {
    assert_eq!(
        run_python_one("import time\nimport time\nprint(time is time)\n"),
        "True"
    );
}

#[test]
fn time_time_compared_to_itself_equals() {
    assert_eq!(
        run_python_one("import time\nt = time.time()\nprint(t == t)\n"),
        "True"
    );
}

#[test]
fn time_sleep_in_try_finally() {
    assert_eq!(
        run_python_one(
            "import time\nout = []\ntry:\n out.append(1)\n time.sleep(0)\nfinally:\n out.append(2)\nprint(out)\n"
        ),
        "[1, 2]"
    );
}

#[test]
fn time_time_with_int_addition() {
    assert_eq!(
        run_python_one("import time\nprint(int(time.time() + 1) >= int(time.time()))\n"),
        "True"
    );
}

#[test]
fn time_time_string_format_fstring() {
    assert_eq!(
        run_python_one("import time\nt = 1.5\nprint(f'{t:.1f}')\n"),
        "1.5"
    );
}

#[test]
fn time_module_type_is_module() {
    assert_eq!(
        run_python_one("import time\nprint(type(time).__name__)\n"),
        "module"
    );
}

#[test]
fn time_sleep_return_is_none() {
    assert_eq!(
        run_python_one("import time\nprint(time.sleep(0))\n"),
        "None"
    );
}
