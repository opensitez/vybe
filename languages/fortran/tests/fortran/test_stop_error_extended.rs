//! Extended STOP / ERROR STOP coverage: message and code forms, guarded branches,
//! and halt-after-print semantics. Distinct from `test_legacy.rs`, `test_fortran2018.rs`,
//! and `test_control_flow_extended.rs`.

use super::helpers::compile_ok;

fortran_cases! {
    // ── STOP with message after print ────────────────────────────────

    stop_message_after_banner => {
        "program t\nprint *, 'banner'\nstop 'halted'\nprint *, 'tail'\nend program t\n",
        ["banner"]
    };

    stop_integer_code_after_ready => {
        "program t\nprint *, 'ready'\nstop 0\nprint *, 'after'\nend program t\n",
        ["ready"]
    };

    diagnostics_then_stop_message => {
        "program t\nprint *, 'step1'\nprint *, 'step2'\nstop 'done'\nprint *, 'step3'\nend program t\n",
        ["step1", "step2"]
    };

    // ── Guarded STOP in IF ───────────────────────────────────────────

    guarded_stop_in_if_not_taken => {
        "program t\nlogical :: ok = .true.\nif (.not. ok) stop 1\nprint *, 'continued'\nend program t\n",
        ["continued"]
    };

    guarded_stop_in_if_taken_halts => {
        "program t\ninteger :: code = 1\nif (code /= 0) stop 0\nprint *, 'run'\nend program t\n",
        []
    };

    guarded_stop_after_warning_print => {
        "program t\ninteger :: n = 2\nprint *, 'warning'\nif (n > 1) stop 1\nprint *, 'tail'\nend program t\n",
        ["warning"]
    };

    // ── ERROR STOP with code (guard not taken) ───────────────────────

    error_stop_code_guard_not_taken => {
        "program t\nlogical :: err = .false.\nprint *, 'check'\nif (err) error stop 7\nprint *, 'ok'\nend program t\n",
        ["check", "ok"]
    };

    stop_in_loop_with_conditional_halt => {
        "program t\ninteger :: i\ninteger :: n\nn = 1\n do i = 1, 3\n    if (i == n) stop 'stop-on-first'\n    print *, i\n end do\n print *, 'tail'\nend program t\n",
        []
    };

    stop_with_code_expression => {
        "program t\ninteger :: code = 2\nif (code > 1) then\n    stop code\nelse\n    print *, 'continue'\nend if\nprint *, 'never'\nend program t\n",
        []
    };

    error_stop_in_nested_control_not_taken => {
        "program t\nlogical :: fail = .false.\nlogical :: inner = .false.\nif (fail) then\n    print *, 'bad'\nelse if (inner) then\n    error stop 'nested bad'\nelse\n    print *, 'ok'\nend if\nprint *, 'done'\nend program t\n",
        ["ok", "done"]
    };

    stop_or_error_goto_like_fallthrough => {
        "program t\ninteger :: code\ncode = 0\nif (code /= 0) stop code\nif (code == 0) then\n    print *, 'ok'\nelse\n    error stop 'bad'\nend if\nprint *, 'done'\nend program t\n",
        ["ok", "done"]
    };
}

// ── Pure STOP / ERROR STOP (compile-only) ─────────────────────────

#[test]
fn stop_message_only() {
    compile_ok("program t\n  stop 'clean exit'\nend program t\n");
}

#[test]
fn stop_integer_literal_no_print() {
    compile_ok("program t\n  stop 0\nend program t\n");
}

#[test]
fn error_stop_integer_code() {
    compile_ok(
        r#"
program t
    logical :: ok = .true.
    if (.not. ok) error stop 1
    print *, 'fine'
end program t
"#,
    );
}

#[test]
fn error_stop_variable_code() {
    compile_ok(
        r#"
program t
    integer :: code = 0
    if (code /= 0) error stop code
    print *, 'ok'
end program t
"#,
    );
}

#[test]
fn error_stop_message_form() {
    compile_ok(
        r#"
program t
    logical :: bad = .false.
    if (bad) error stop 'fatal condition'
    print *, 'clear'
end program t
"#,
    );
}

#[test]
fn error_stop_with_quiet_clause() {
    compile_ok(
        r#"
program t
    logical :: bad = .false.
    if (bad) error stop 9, quiet = .true.
    print *, 'ok'
end program t
"#,
    );
}
