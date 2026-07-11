use super::helpers::compile_ok;
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
