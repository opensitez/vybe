use super::helpers::run_prints;

#[test]
fn program_unit_startup_states_module_constants_seed_program_order() {
    let out = run_prints(
        r#"
module startup_constants
    integer, parameter :: base = 3
    integer :: value
    integer :: computed
    value = base + 7
    computed = value * 2
contains
    integer function get()
        get = value + computed
    end function get
end module startup_constants

program program_unit_startup_states_module_constants_seed_program_order
    use startup_constants
    print *, base
    print *, value
    print *, computed
    print *, get()
end program program_unit_startup_states_module_constants_seed_program_order
"#,
    );
    assert_eq!(out, vec!["3", "10", "20", "30"]);
}

#[test]
fn program_unit_startup_states_derived_type_default_init() {
    let out = run_prints(
        r#"
module startup_types
    type cfg
        integer :: a = 1
        integer :: b = 2
        integer :: c = a + b
    end type cfg
end module startup_types

program program_unit_startup_states_derived_type_default_init
    use startup_types
    type(cfg) :: item
    print *, item%a
    print *, item%b
    print *, item%c
end program program_unit_startup_states_derived_type_default_init
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn program_unit_startup_states_data_initialization_before_block() {
    let out = run_prints(
        r#"
program program_unit_startup_states_data_initialization_before_block
    integer :: seed
    data seed /99/
    block
        integer :: doubled
        doubled = seed * 2
        print *, seed
        print *, doubled
    end block
end program program_unit_startup_states_data_initialization_before_block
"#,
    );
    assert_eq!(out, vec!["99", "198"]);
}

#[test]
fn program_unit_startup_states_block_data_like_seed() {
    let out = run_prints(
        r#"
module block_state
    integer :: a
    integer :: b
    integer, save :: token = 11
    contains
        subroutine init()
            a = 5
            b = a + token
        end subroutine init
end module block_state

program program_unit_startup_states_block_data_like_seed
    use block_state
    call init()
    print *, a
    print *, b
    print *, token
end program program_unit_startup_states_block_data_like_seed
"#,
    );
    assert_eq!(out, vec!["5", "16", "11"]);
}

#[test]
fn program_unit_startup_states_nested_modules_init_order() {
    let out = run_prints(
        r#"
module alpha
    integer, parameter :: p = 2
end module alpha

module beta
    use alpha
    integer, parameter :: q = p + 3
end module beta

program program_unit_startup_states_nested_modules_init_order
    use beta
    print *, p
    print *, q
end program program_unit_startup_states_nested_modules_init_order
"#,
    );
    assert_eq!(out, vec!["2", "5"]);
}

#[test]
fn program_unit_startup_states_common_init_before_subroutine_use() {
    let out = run_prints(
        r#"
module startup_mod
    integer :: counter = 0
contains
    subroutine bump()
        counter = counter + 1
    end subroutine bump
end module startup_mod

program program_unit_startup_states_common_init_before_subroutine_use
    use startup_mod
    call bump()
    call bump()
    print *, counter
end program program_unit_startup_states_common_init_before_subroutine_use
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn program_unit_startup_states_program_unit_sequence_points() {
    let out = run_prints(
        r#"
program program_unit_startup_states_program_unit_sequence_points
    integer :: x = 1
    integer :: y = 2
    integer :: z
    z = x + y
    print *, x
    print *, y
    print *, z
end program program_unit_startup_states_program_unit_sequence_points
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn program_unit_startup_states_initialization_dependencies() {
    let out = run_prints(
        r#"
program program_unit_startup_states_initialization_dependencies
    integer :: first = 4
    integer :: second = first + 1
    integer :: third = second + 1
    print *, first
    print *, second
    print *, third
end program program_unit_startup_states_initialization_dependencies
"#,
    );
    assert_eq!(out, vec!["4", "5", "6"]);
}

#[test]
fn program_unit_startup_states_block_rebind_seed() {
    let out = run_prints(
        r#"
program program_unit_startup_states_block_rebind_seed
    integer :: count
    count = 1
    block
        integer :: count
        count = 99
        print *, count
    end block
    print *, count
end program program_unit_startup_states_block_rebind_seed
"#,
    );
    assert_eq!(out, vec!["99", "1"]);
}
