use super::helpers::{compile_ok, run_prints};

#[test]
fn stop_error_status_signaling_simple_stop_form() {
    compile_ok(
        "program stop_error_status_signaling_simple_stop_form\n\
            integer :: x\n\
            x = 1\n\
            if (x > 0) stop 0\n\
            print *, 'unreachable'\n\
        end program stop_error_status_signaling_simple_stop_form\n",
    );
}

#[test]
fn stop_error_status_signaling_simple_stop_form_runtime() {
    let out = run_prints(
        "program stop_error_status_signaling_simple_stop_form_runtime\n\
integer :: x\nx = 0\nif (x > 0) stop 0\nprint *, x\n\
end program stop_error_status_signaling_simple_stop_form_runtime\n",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stop_error_status_signaling_error_stop_with_code() {
    compile_ok(
        "program stop_error_status_signaling_error_stop_with_code\n\
            integer :: x\n\
            x = 0\n\
            if (x /= 0) error stop 17\n\
            print *, x\n\
        end program stop_error_status_signaling_error_stop_with_code\n",
    );
}

#[test]
fn stop_error_status_signaling_message_form() {
    compile_ok(
        "program stop_error_status_signaling_message_form\n\
            logical :: failed\n\
            failed = .false.\n\
            if (failed) stop 'all-good'\n\
            print *, 'running'\n\
        end program stop_error_status_signaling_message_form\n",
    );
}

#[test]
fn stop_error_status_signaling_error_stop_message_form() {
    compile_ok(
        "program stop_error_status_signaling_error_stop_message_form\n\
            integer :: n\n\
            n = 1\n\
            if (n < 0) error stop 'invalid-negative'\n\
            print *, n\n\
        end program stop_error_status_signaling_error_stop_message_form\n",
    );
}

#[test]
fn stop_error_status_signaling_nested_stop_paths() {
    compile_ok(
        "program stop_error_status_signaling_nested_stop_paths\n\
            integer :: i\n\
            do i = 1, 2\n\
                if (i == 3) stop 1\n\
            end do\n\
            print *, 'done'\n\
        end program stop_error_status_signaling_nested_stop_paths\n",
    );
}

#[test]
fn stop_error_status_signaling_error_stop_in_subroutine() {
    compile_ok(
        "program stop_error_status_signaling_error_stop_in_subroutine\n\
            call guard(.true.)\n\
            print *, 'after'\n\
        contains\n            subroutine guard(flag)\n                logical, intent(in) :: flag\n                if (flag) error stop 2\n            end subroutine guard\n        end program stop_error_status_signaling_error_stop_in_subroutine\n",
    );
}

#[test]
fn stop_error_status_signaling_error_stop_in_subroutine_runtime_guarded() {
    let out = run_prints(
        "program stop_error_status_signaling_error_stop_in_subroutine_runtime_guarded\n\
    call guard(.false.)\n\
    print *, 'after'\n\
contains\n+    subroutine guard(flag)\n+        logical, intent(in) :: flag\n+        if (flag) error stop 2\n+    end subroutine guard\n+end program stop_error_status_signaling_error_stop_in_subroutine_runtime_guarded\n",
    );
    assert_eq!(out, vec!["after"]);
}

#[test]
fn stop_error_status_signaling_stop_with_zero_code_is_terminal() {
    compile_ok(
        "program stop_error_status_signaling_stop_with_zero_code_is_terminal\n\
            if (.false.) stop 0\n\
            print *, 'ok'\n\
        end program stop_error_status_signaling_stop_with_zero_code_is_terminal\n",
    );
}

#[test]
fn stop_error_status_signaling_error_stop_with_guard_variable() {
    compile_ok(
        "program stop_error_status_signaling_error_stop_with_guard_variable\n\
            logical :: do_stop\n\
            do_stop = .false.\n\
            if (do_stop) error stop 99\n\
            print *, 1\n\
        end program stop_error_status_signaling_error_stop_with_guard_variable\n",
    );
}

#[test]
fn stop_error_status_signaling_error_stop_message_with_quiet_false() {
    compile_ok(
        "program stop_error_status_signaling_error_stop_message_with_quiet_false\n\
            integer :: flag\n\
            flag = 0\n\
            if (flag == 0) then\n\
                error stop 'status ok', quiet = .false.\n\
            else\n\
                print *, 'not triggered'\n\
            end if\n\
            print *, 'end'\n\
        end program stop_error_status_signaling_error_stop_message_with_quiet_false\n",
    );
}

#[test]
fn stop_error_status_signaling_error_stop_with_identifier_code() {
    compile_ok(
        "program stop_error_status_signaling_error_stop_with_identifier_code\n\
            integer :: status\n\
            status = 3\n\
            if (status > 0) error stop status\n\
            print *, status\n\
        end program stop_error_status_signaling_error_stop_with_identifier_code\n",
    );
}
