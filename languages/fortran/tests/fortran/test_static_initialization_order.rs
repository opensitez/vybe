use super::helpers::run_prints;

#[test]
fn static_initialization_order_dependency_chain() {
    let out = run_prints(
        r#"
program static_initialization_order_dependency_chain
    integer :: a = 1
    integer :: b = a + 2
    integer :: c = b * 3
    print *, a
    print *, b
    print *, c
end program static_initialization_order_dependency_chain
"#,
    );
    assert_eq!(out, vec!["1", "3", "9"]);
}

#[test]
fn static_initialization_order_module_and_program_order() {
    let out = run_prints(
        r#"
module m
    integer, parameter :: base = 3
    integer, save :: offset = base + 1
    integer, save :: total = offset * 4
end module m

program static_initialization_order_module_and_program_order
    use m
    print *, total
    print *, base
    print *, offset
end program static_initialization_order_module_and_program_order
"#,
    );
    assert_eq!(out, vec!["16", "3", "4"]);
}

#[test]
fn static_initialization_order_derived_array_defaults() {
    let out = run_prints(
        r#"
program static_initialization_order_derived_array_defaults
    integer :: table(3) = (/1, 2, 3/)
    integer :: total
    total = sum(table)
    print *, table(1)
    print *, table(3)
    print *, total
end program static_initialization_order_derived_array_defaults
"#,
    );
    assert_eq!(out, vec!["1", "3", "6"]);
}

#[test]
fn static_initialization_order_character_default_len() {
    let out = run_prints(
        r#"
program static_initialization_order_character_default_len
    character(len=4) :: token = 'seed'
    character(len=len_trim(token)) :: short
    short = token
    print *, trim(short)
    print *, len_trim(short)
end program static_initialization_order_character_default_len
"#,
    );
    assert_eq!(out, vec!["seed", "4"]);
}

#[test]
fn static_initialization_order_function_use_of_parameter() {
    let out = run_prints(
        r#"
program static_initialization_order_function_use_of_parameter
    integer, parameter :: p = 7
    integer :: q = p
    print *, adjustl(int2str(q))
contains
    character(len=16) function int2str(v)
        integer, intent(in) :: v
        write(int2str, '(I0)') v
    end function int2str
end program static_initialization_order_function_use_of_parameter
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn static_initialization_order_block_data_like_save() {
    let out = run_prints(
        r#"
program static_initialization_order_block_data_like_save
    integer, save :: counter
    counter = 5
    print *, counter
    counter = counter + 1
    print *, counter
end program static_initialization_order_block_data_like_save
"#,
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn static_initialization_order_component_with_intent_like_defaults() {
    let out = run_prints(
        r#"
program static_initialization_order_component_with_intent_like_defaults
    type state
        integer :: left = 1
        integer :: right = 2
        integer :: total
    end type state
    type(state) :: s
    s%total = s%left + s%right
    print *, s%left
    print *, s%right
    print *, s%total
end program static_initialization_order_component_with_intent_like_defaults
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn static_initialization_order_order_inside_block() {
    let out = run_prints(
        r#"
program static_initialization_order_order_inside_block
    integer :: base = 10
    block
        integer :: nested = base + 1
        print *, nested
    end block
    print *, base
end program static_initialization_order_order_inside_block
"#,
    );
    assert_eq!(out, vec!["11", "10"]);
}

#[test]
fn static_initialization_order_multiple_saves() {
    let out = run_prints(
        r#"
program static_initialization_order_multiple_saves
    integer, save :: a = 1
    integer, save :: b = a + 1
    integer, save :: c = b + 1
    print *, a
    print *, b
    print *, c
end program static_initialization_order_multiple_saves
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn static_initialization_order_parameterized_save_dependency() {
    let out = run_prints(
        r#"
module static_init_param_dependency
    integer, parameter :: scale = 2
    integer, save :: base = scale
    integer, save :: doubled = base * scale
end module static_init_param_dependency

program static_initialization_order_parameterized_save_dependency
    use static_init_param_dependency
    print *, base
    print *, doubled
end program static_initialization_order_parameterized_save_dependency
"#,
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn static_initialization_order_array_constructor_default() {
    let out = run_prints(
        r#"
program static_initialization_order_array_constructor_default
    integer, save :: values(4) = (/1, 2, 3, 4/)
    integer :: i
    i = values(1) + values(4)
    print *, values(2)
    print *, i
end program static_initialization_order_array_constructor_default
"#,
    );
    assert_eq!(out, vec!["2", "5"]);
}

#[test]
fn static_initialization_order_character_len_from_constant() {
    let out = run_prints(
        r#"
program static_initialization_order_character_len_from_constant
    integer, parameter :: name_len = 5
    character(len=name_len) :: tag = 'abc'
    print *, len(tag)
    print *, trim(tag)
end program static_initialization_order_character_len_from_constant
"#,
    );
    assert_eq!(out, vec!["5", "abc"]);
}
