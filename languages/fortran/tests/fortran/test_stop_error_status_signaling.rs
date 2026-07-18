use super::helpers::compile_ok;

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
