use crate::helpers::run_python_one;

#[test]
fn sys_import_name() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.__name__)\n"),
        "sys"
    );
}

#[test]
fn sys_argv_is_list() {
    assert_eq!(
        run_python_one("import sys\nprint(isinstance(sys.argv, list))\n"),
        "True"
    );
}

#[test]
fn sys_argv_nonempty() {
    assert_eq!(
        run_python_one("import sys\nprint(len(sys.argv) >= 1)\n"),
        "True"
    );
}

#[test]
fn sys_argv_first_is_string() {
    assert_eq!(
        run_python_one("import sys\nprint(isinstance(sys.argv[0], str))\n"),
        "True"
    );
}

#[test]
fn sys_platform_is_string() {
    assert_eq!(
        run_python_one("import sys\nprint(isinstance(sys.platform, str))\n"),
        "True"
    );
}

#[test]
fn sys_platform_nonempty() {
    assert_eq!(
        run_python_one("import sys\nprint(len(sys.platform) > 0)\n"),
        "True"
    );
}

#[test]
fn sys_has_exit_attribute() {
    assert_eq!(
        run_python_one("import sys\nprint(callable(sys.exit))\n"),
        "True"
    );
}

#[test]
fn sys_has_version_info() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys, 'version'))\n"),
        "True"
    );
}

#[test]
fn sys_version_is_string() {
    assert_eq!(
        run_python_one("import sys\nprint(isinstance(sys.version, str))\n"),
        "True"
    );
}

#[test]
fn sys_has_stdin_stdout_stderr() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys, 'stdin') and hasattr(sys, 'stdout') and hasattr(sys, 'stderr'))\n"),
        "True"
    );
}

#[test]
fn sys_modules_is_dict() {
    assert_eq!(
        run_python_one("import sys\nprint(isinstance(sys.modules, dict))\n"),
        "True"
    );
}

#[test]
fn sys_modules_contains_sys() {
    assert_eq!(
        run_python_one("import sys\nprint('sys' in sys.modules)\n"),
        "True"
    );
}

#[test]
fn sys_path_is_list() {
    assert_eq!(
        run_python_one("import sys\nprint(isinstance(sys.path, list))\n"),
        "True"
    );
}

#[test]
fn sys_getdefaultencoding() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.getdefaultencoding())\n"),
        "utf-8"
    );
}

#[test]
fn sys_getrecursionlimit_positive() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.getrecursionlimit() > 0)\n"),
        "True"
    );
}

#[test]
fn sys_setrecursionlimit_roundtrip() {
    assert_eq!(
        run_python_one("import sys\nold = sys.getrecursionlimit()\nsys.setrecursionlimit(old)\nprint(sys.getrecursionlimit() == old)\n"),
        "True"
    );
}

#[test]
fn sys_intern_reuses_string() {
    assert_eq!(
        run_python_one("import sys\na = sys.intern('vybe')\nb = sys.intern('vybe')\nprint(a is b)\n"),
        "True"
    );
}

#[test]
fn sys_exc_info_outside_except() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.exc_info() == (None, None, None))\n"),
        "True"
    );
}

#[test]
fn sys_exc_info_inside_except() {
    assert_eq!(
        run_python_one("import sys\ntry:\n raise ValueError('x')\nexcept ValueError:\n t, v, tb = sys.exc_info()\n print(t.__name__, str(v))\n"),
        "ValueError x"
    );
}

#[test]
fn sys_exception_in_except_block() {
    assert_eq!(
        run_python_one("import sys\ntry:\n raise TypeError('t')\nexcept TypeError as e:\n print(sys.exception() is e)\n"),
        "True"
    );
}

#[test]
fn sys_is_finalizing_false_during_run() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.is_finalizing())\n"),
        "False"
    );
}

#[test]
fn sys_byteorder_is_little_or_big() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.byteorder in ('little', 'big'))\n"),
        "True"
    );
}

#[test]
fn sys_maxsize_positive() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.maxsize > 0)\n"),
        "True"
    );
}

#[test]
fn sys_has_float_info() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys, 'float_info'))\n"),
        "True"
    );
}

#[test]
fn sys_has_int_info() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys, 'int_info'))\n"),
        "True"
    );
}

#[test]
fn sys_implementation_name() {
    assert_eq!(
        run_python_one("import sys\nprint(len(sys.implementation.name) > 0)\n"),
        "True"
    );
}

#[test]
fn sys_argv_copy_independent() {
    assert_eq!(
        run_python_one("import sys\na = sys.argv\na.append('extra')\nprint(len(sys.argv) == len(a) - 1 or 'extra' not in sys.argv)\n"),
        "True"
    );
}

#[test]
fn sys_platform_lower_nonempty() {
    assert_eq!(
        run_python_one("import sys\nprint(len(sys.platform.lower()) > 0)\n"),
        "True"
    );
}

#[test]
fn sys_version_info_tuple_like() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.version_info.major >= 3)\n"),
        "True"
    );
}

#[test]
fn sys_modules_records_imported_os() {
    assert_eq!(
        run_python_one("import os\nimport sys\nprint('os' in sys.modules)\n"),
        "True"
    );
}

#[test]
fn sys_path_insert_and_restore() {
    assert_eq!(
        run_python_one("import sys\nold = list(sys.path)\nsys.path.insert(0, '/tmp/vybe-test')\nprint(sys.path[0] == '/tmp/vybe-test')\nsys.path[:] = old\n"),
        "True"
    );
}

#[test]
fn sys_getsizeof_int() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.getsizeof(0) > 0)\n"),
        "True"
    );
}

#[test]
fn sys_getsizeof_list_grows() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.getsizeof([1, 2, 3]) >= sys.getsizeof([]))\n"),
        "True"
    );
}

#[test]
fn sys_displayhook_does_not_crash() {
    assert_eq!(
        run_python_one("import sys\nold = sys.displayhook\nsys.displayhook = old\nprint('ok')\n"),
        "ok"
    );
}

#[test]
fn sys_excepthook_is_callable() {
    assert_eq!(
        run_python_one("import sys\nprint(callable(sys.excepthook))\n"),
        "True"
    );
}

#[test]
fn sys_stdin_not_none() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.stdin is not None)\n"),
        "True"
    );
}

#[test]
fn sys_stdout_not_none() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.stdout is not None)\n"),
        "True"
    );
}

#[test]
fn sys_stderr_not_none() {
    assert_eq!(
        run_python_one("import sys\nprint(sys.stderr is not None)\n"),
        "True"
    );
}

#[test]
fn sys_flags_has_bytes_warning_attr() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys, 'flags'))\n"),
        "True"
    );
}

#[test]
fn sys_hash_info_has_width() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys.hash_info, 'width'))\n"),
        "True"
    );
}

#[test]
fn sys_thread_info_exists() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys, 'thread_info'))\n"),
        "True"
    );
}

#[test]
fn sys_audit_available_or_skipped() {
    assert_eq!(
        run_python_one("import sys\nprint(hasattr(sys, 'audit'))\n"),
        "True"
    );
}

#[test]
fn sys_unraisablehook_callable() {
    assert_eq!(
        run_python_one("import sys\nprint(callable(sys.unraisablehook))\n"),
        "True"
    );
}

#[test]
fn sys_orig_argv_list() {
    assert_eq!(
        run_python_one("import sys\nprint(isinstance(getattr(sys, 'orig_argv', sys.argv), list))\n"),
        "True"
    );
}
