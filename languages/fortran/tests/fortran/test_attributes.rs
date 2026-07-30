use super::helpers::{compile_ok, run_prints};
macro_rules! c {
    ($name:ident,$src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}
c!(
    attr_alloc_01,
    "program p\ninteger, allocatable :: a(:)\nend program p\n"
);
c!(
    attr_ptr_02,
    "program p\ninteger, pointer :: p\nend program p\n"
);
c!(
    attr_target_03,
    "program p\ninteger, target :: x\nend program p\n"
);
c!(
    attr_save_04,
    "program p\ninteger, save :: x\nend program p\n"
);
c!(
    attr_protected_05,
    "module m\ninteger, protected :: x\nend module m\n"
);
c!(
    attr_volatile_06,
    "program p\ninteger, volatile :: x\nend program p\n"
);
c!(
    attr_async_07,
    "program p\ninteger, asynchronous :: x\nend program p\n"
);
c!(
    attr_bindc_08,
    "type, bind(c) :: t\ninteger :: x\nend type t\n"
);
c!(
    attr_sequence_09,
    "type, sequence :: t\ninteger :: x\nend type t\n"
);
c!(
    attr_extends_10,
    "type :: b\ninteger::x\nend type b\ntype, extends(b) :: c\ninteger::y\nend type c\n"
);
c!(
    attr_abstract_11,
    "type, abstract :: t\ninteger::x\nend type t\n"
);
c!(
    attr_deferred_12,
    "type, abstract :: t\ncontains\nprocedure(p),deferred::s\nend type t\nabstract interface\nsubroutine p(this)\nimport t\nclass(t)::this\nend\nend interface\n"
);
c!(
    attr_non_over_13,
    "module m\ntype::t\ncontains\nprocedure,non_overridable::s\nend type\ncontains\nsubroutine s(this)\nclass(t)::this\nend\nend module m\n"
);
c!(
    attr_private_14,
    "module m\nprivate\ninteger::x\nend module m\n"
);
c!(
    attr_public_15,
    "module m\npublic :: x\ninteger::x\nend module m\n"
);
c!(
    attr_parameter_16,
    "program p\ninteger, parameter :: n=4\nprint *, n\nend program p\n"
);
c!(
    attr_intent_17,
    "subroutine s(x)\ninteger,intent(inout)::x\nend subroutine s\n"
);
c!(
    attr_optional_18,
    "subroutine s(x)\ninteger,optional::x\nend subroutine s\n"
);
c!(
    attr_value_19,
    "subroutine s(x)\ninteger,value::x\nend subroutine s\n"
);
c!(attr_codim_20, "program p\ninteger :: x[*]\nend program p\n");
c!(
    attr_dimension_21,
    "program p\ninteger, dimension(3) :: a\nend program p\n"
);
c!(
    attr_contiguous_22,
    "subroutine s(a)\nreal,contiguous::a(:)\nend subroutine s\n"
);
c!(attr_external_23, "program p\nexternal f\nend program p\n");
c!(
    attr_intrinsic_24,
    "program p\nintrinsic abs\nprint *, abs(-1)\nend program p\n"
);
c!(
    attr_pointer_comp_25,
    "type :: t\ninteger, pointer :: p\nend type t\n"
);
c!(
    attr_alloc_comp_26,
    "type :: t\ninteger, allocatable :: a(:)\nend type t\n"
);
c!(
    attr_len_char_27,
    "program p\ncharacter(len=8) :: s\nend program p\n"
);
c!(
    attr_deferred_len_28,
    "program p\ncharacter(len=:), allocatable :: s\nend program p\n"
);
c!(
    attr_class_star_29,
    "subroutine s(x)\nclass(*) :: x\nend subroutine s\n"
);
c!(
    attr_assumed_rank_30,
    "subroutine s(a)\ninteger :: a(..)\nend subroutine s\n"
);

#[test]
fn attr_allocatable_runtime_value_flow() {
    let out = run_prints(
        r#"
program attr_allocatable_runtime_value_flow
    integer, allocatable :: values(:)
    allocate(values(2))
    values = (/10, 20/)
    print *, values(1)
    print *, values(2)
end program attr_allocatable_runtime_value_flow
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn attr_pointer_runtime_aliasing() {
    let out = run_prints(
        r#"
program attr_pointer_runtime_aliasing
    integer, target :: storage
    integer, pointer :: p
    storage = 33
    p => storage
    print *, p
    p = 44
    print *, storage
end program attr_pointer_runtime_aliasing
"#,
    );
    assert_eq!(out, vec!["33", "44"]);
}

#[test]
fn attr_save_runtime_default_stable() {
    let out = run_prints(
        r#"
program attr_save_runtime_default_stable
    integer, save :: counter = 0
    counter = counter + 1
    print *, counter
    counter = counter + 1
    print *, counter
end program attr_save_runtime_default_stable
"#,
);
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn attr_protected_module_accessor() {
    let out = run_prints(
        r#"
module protected_access_mod
    integer, protected :: value = 9
contains
    integer function get_value()
        get_value = value
    end function get_value
end module protected_access_mod

program attr_protected_module_accessor
    use protected_access_mod
    print *, get_value()
end program attr_protected_module_accessor
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn attr_sequence_and_extends_values() {
    let out = run_prints(
        r#"
program attr_sequence_and_extends_values
    type :: Base
        integer :: base_field = 5
    end type Base

    type, extends(Base) :: Child
        integer :: child_field = 7
    end type Child

    type(Child) :: c
    print *, c%base_field
    print *, c%child_field
    print *, c%base_field + c%child_field
end program attr_sequence_and_extends_values
"#,
    );
    assert_eq!(out, vec!["5", "7", "12"]);
}

#[test]
fn attr_intent_optional_value_flow() {
    let out = run_prints(
        r#"
program attr_intent_optional_value_flow
    call with_attributes(10, .true.)
contains
    subroutine with_attributes(x, flag)
        integer, optional, intent(in) :: x
        logical, intent(in) :: flag
        print *, present(x)
        print *, flag
    end subroutine with_attributes
end program attr_intent_optional_value_flow
"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn attr_intent_optional_can_be_absent() {
    let out = run_prints(
        r#"
program attr_intent_optional_can_be_absent
    call with_optional_arg()
    call with_optional_arg(4)
contains
    subroutine with_optional_arg(x)
        integer, optional, intent(inout) :: x
        if (present(x)) then
            print *, x
        else
            print *, -1
        end if
    end subroutine with_optional_arg
end program attr_intent_optional_can_be_absent
"#,
    );
    assert_eq!(out, vec!["-1", "4"]);
}

#[test]
fn attr_parameter_is_compile_time_constant_used_at_runtime() {
    let out = run_prints(
        r#"
program attr_parameter_is_compile_time_constant_used_at_runtime
    integer, parameter :: n = 4
    integer :: a(n)
    a = [1, 2, 3, 4]
    print *, a(n)
end program attr_parameter_is_compile_time_constant_used_at_runtime
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn attr_dimension_declared_array_behaves_like_fixed_shape() {
    let out = run_prints(
        r#"
program attr_dimension_declared_array_behaves_like_fixed_shape
    integer, dimension(3) :: a
    a = [5, 6, 7]
    print *, a(1) + a(2) + a(3)
end program attr_dimension_declared_array_behaves_like_fixed_shape
"#,
    );
    assert_eq!(out, vec!["18"]);
}
