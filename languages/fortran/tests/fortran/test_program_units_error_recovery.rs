use super::helpers::parse_ok;

#[test]
fn program_units_error_recovery_invalid_unit_boundary_fails_parse() {
    assert!(!parse_ok(
        "program program_units_error_recovery_invalid_unit_boundary_fails_parse\n\
            integer :: x\n\
            end\n",
    ));
}

#[test]
fn program_units_error_recovery_invalid_fixed_loop_rejected() {
    assert!(!parse_ok(
        "program program_units_error_recovery_invalid_fixed_loop_rejected\n\
            do i = 1, 10\n\
            print *, i\n\
        end do\n",
    ));
}

#[test]
fn program_units_error_recovery_unbalanced_if_rejected() {
    assert!(!parse_ok(
        "program program_units_error_recovery_unbalanced_if_rejected\n\
            integer :: x\n\
            if (.true.) print *, 1\n\
        end program\n",
    ));
}

#[test]
fn program_units_error_recovery_invalid_type_decl_rejected() {
    assert!(!parse_ok(
        "module program_units_error_recovery_invalid_type_decl_rejected\n\
            type :: item\n\
                integer :: a =\n\
            end type\n\
        end module\n",
    ));
}

#[test]
fn program_units_error_recovery_recovery_after_failure_still_parses_valid_unit() {
    assert!(!parse_ok(
        "program bad\n\
            integer :: x =\n\
        end program\n",
    ));
    assert!(parse_ok(
        "program program_units_error_recovery_recovery_after_failure_still_parses_valid_unit\n\
            print *, 42\n\
        end program\n",
    ));
}

#[test]
fn program_units_error_recovery_mismatched_labels_rejected() {
    assert!(!parse_ok(
        "program p\n\
        10 print *, 1\n\
        20 print *, 2\n\
        go to 30\n\
        end program\n",
    ));
}

#[test]
fn program_units_error_recovery_invalid_return_name() {
    assert!(!parse_ok(
        "function f() result(r)\n\
            r = 1\n\
        end f\n",
    ));
}

#[test]
fn program_units_error_recovery_invalid_keyword_sequence() {
    assert!(!parse_ok(
        "program p\n\
            if (.true.) then\n\
                print *, 1\n\
            else\n\
                print *, 2\n\
        end program p\n",
    ));
}
