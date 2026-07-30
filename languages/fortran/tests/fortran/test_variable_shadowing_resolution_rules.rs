use super::helpers::run_prints;

#[test]
fn variable_shadowing_resolution_rules_block_local_hides_host() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_block_local_hides_host
    integer :: x
    x = 1
    block
        integer :: x
        x = 2
        print *, x
    end block
    print *, x
end program variable_shadowing_resolution_rules_block_local_hides_host
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn variable_shadowing_resolution_rules_nested_blocks() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_nested_blocks
    integer :: value
    value = 10
    block
        integer :: value
        value = 20
        block
            integer :: value
            value = 30
            print *, value
        end block
        print *, value
    end block
    print *, value
end program variable_shadowing_resolution_rules_nested_blocks
"#,
    );
    assert_eq!(out, vec!["30", "20", "10"]);
}

#[test]
fn variable_shadowing_resolution_rules_module_procedure_scope() {
    let out = run_prints(
        r#"
module scope_rules_mod
    integer :: token = 1
contains
    subroutine report()
        integer :: token
        token = 4
        print *, token
    end subroutine report
end module scope_rules_mod

program variable_shadowing_resolution_rules_module_procedure_scope
    use scope_rules_mod
    print *, token
    call report()
    print *, token
end program variable_shadowing_resolution_rules_module_procedure_scope
"#,
    );
    assert_eq!(out, vec!["1", "4", "1"]);
}

#[test]
fn variable_shadowing_resolution_rules_interface_shadows() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_interface_shadows
    integer :: limit
    limit = 3
    print *, wrap(limit)
contains
    integer function wrap(value)
        integer, intent(in) :: value
        integer :: limit
        limit = value + 1
        wrap = limit
    end function wrap
end program variable_shadowing_resolution_rules_interface_shadows
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn variable_shadowing_resolution_rules_procedure_argument_priority() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_procedure_argument_priority
    integer :: x
    x = 8
    print *, compute(x)
contains
    integer function compute(x)
        integer, intent(in) :: x
        integer :: local
        local = x + 2
        compute = local
    end function compute
end program variable_shadowing_resolution_rules_procedure_argument_priority
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn variable_shadowing_resolution_rules_name_hides_intrinsic() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_name_hides_intrinsic
    integer :: sum
    sum = 5
    print *, bump(sum)
contains
    integer function bump(sum)
        integer, intent(in) :: sum
        bump = sum + 1
    end function bump
end program variable_shadowing_resolution_rules_name_hides_intrinsic
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn variable_shadowing_resolution_rules_type_component_hiding() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_type_component_hiding
    type item
        integer :: value = 1
    end type item
    type(item) :: it
    integer :: value
    value = 9
    it%value = 3
    print *, value
    print *, it%value
end program variable_shadowing_resolution_rules_type_component_hiding
"#,
    );
    assert_eq!(out, vec!["9", "3"]);
}

#[test]
fn variable_shadowing_resolution_rules_imported_name_masking() {
    let out = run_prints(
        r#"
module shadow_a
    integer :: token = 1
end module shadow_a

program variable_shadowing_resolution_rules_imported_name_masking
    use shadow_a, only: module_token => token
    integer :: token
    token = 9
    print *, token
    print *, module_token
end program variable_shadowing_resolution_rules_imported_name_masking
"#,
    );
    assert_eq!(out, vec!["9", "1"]);
}

#[test]
fn variable_shadowing_resolution_rules_do_index_shadowing() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_do_index_shadowing
    integer :: i
    i = 1
    do i = 1, 2
        print *, i
    end do
    print *, i
end program variable_shadowing_resolution_rules_do_index_shadowing
"#,
    );
    assert_eq!(out, vec!["1", "2", "2"]);
}

#[test]
fn variable_shadowing_resolution_rules_named_block_scope() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_named_block_scope
    integer :: tally
    tally = 1
    block named_block
        integer :: tally
        tally = 10
        print *, tally
    end block named_block
    print *, tally
end program variable_shadowing_resolution_rules_named_block_scope
"#,
    );
    assert_eq!(out, vec!["10", "1"]);
}

#[test]
fn variable_shadowing_resolution_rules_if_block_scope() {
    let out = run_prints(
        r#"
program variable_shadowing_resolution_rules_if_block_scope
    integer :: level
    level = 3
    if (level > 0) then
        integer :: level
        level = 11
        print *, level
    end if
    print *, level
end program variable_shadowing_resolution_rules_if_block_scope
"#,
    );
    assert_eq!(out, vec!["11", "3"]);
}

#[test]
fn variable_shadowing_resolution_rules_module_procedure_argument_same_name() {
    let out = run_prints(
        r#"
module shadow_args_mod
    integer :: value = 13
contains
    subroutine emit(value)
        integer, intent(in) :: value
        print *, value
    end subroutine emit
end module shadow_args_mod

program variable_shadowing_resolution_rules_module_procedure_argument_same_name
    use shadow_args_mod
    call emit(99)
    print *, value
end program variable_shadowing_resolution_rules_module_procedure_argument_same_name
"#,
    );
    assert_eq!(out, vec!["99", "13"]);
}
