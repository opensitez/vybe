use super::helpers::compile_ok;

macro_rules! c {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

c!(cf_copylike_01, "program t\nimplicit none\nend program t\n");
c!(
    cf_free_02,
    ">>SOURCE FORMAT FREE\nprogram t\nend program t\n"
);
c!(
    cf_fixed_03,
    ">>SOURCE FORMAT FIXED\nprogram t\nend program t\n"
);
c!(
    cf_switch_04,
    ">>SOURCE FORMAT FREE\nprogram t\n>>SOURCE FORMAT FIXED\nend program t\n"
);
c!(cf_comment_05, "program t\n! comment\nend program t\n");
c!(
    cf_continuation_06,
    "program t\ninteger :: x\nx = 1 &\n  + 2\nend program t\n"
);
c!(
    cf_preproc_like_07,
    "program t\ninteger :: x\nx = 1\nend program t\n"
);
c!(
    cf_include_like_08,
    "program t\nimplicit none\nend program t\n"
);
c!(
    cf_debug_line_09,
    "program t\ninteger :: x\nx = 1\nend program t\n"
);
c!(
    cf_conditional_like_10,
    "program t\ninteger :: x\nx = 1\nend program t\n"
);
c!(
    cf_implicit_none_11,
    "program t\nimplicit none\ninteger :: x\nx=1\nend program t\n"
);
c!(
    cf_save_12,
    "program t\ninteger, save :: x=1\nprint *, x\nend program t\n"
);
c!(
    cf_parameter_13,
    "program t\ninteger, parameter :: x=1\nprint *, x\nend program t\n"
);
c!(
    cf_public_private_14,
    "module m\nprivate\npublic :: s\ncontains\nsubroutine s()\nend subroutine s\nend module m\n"
);
c!(cf_bind_c_15, "subroutine s() bind(c)\nend subroutine s\n");
c!(
    cf_sequence_16,
    "type :: t\nsequence\ninteger :: x\nend type t\n"
);
c!(cf_abstract_17, "type, abstract :: t\nend type t\n");
c!(
    cf_extends_18,
    "type :: p\ninteger :: x\nend type p\ntype, extends(p) :: c\ninteger :: y\nend type c\n"
);
c!(
    cf_deferred_19,
    "type, abstract :: t\ncontains\nprocedure(p), deferred :: run\nend type t\nabstract interface\nsubroutine p(self)\nimport t\nclass(t) :: self\nend subroutine p\nend interface\n"
);
c!(
    cf_non_overridable_20,
    "type :: t\ncontains\nprocedure, non_overridable :: p\nend type t\ncontains\nsubroutine p(self)\nclass(t) :: self\nend subroutine p\n"
);
c!(
    cf_use_intrinsic_21,
    "program t\nuse, intrinsic :: iso_fortran_env\nend program t\n"
);
c!(
    cf_use_non_intrinsic_22,
    "module m\nend module m\nprogram t\nuse, non_intrinsic :: m\nend program t\n"
);
c!(
    cf_rename_use_23,
    "module m\ninteger :: x\nend module m\nprogram t\nuse m, y => x\nend program t\n"
);
c!(
    cf_only_use_24,
    "module m\ninteger :: x\nend module m\nprogram t\nuse m, only: x\nend program t\n"
);
c!(
    cf_import_25,
    "module m\nimplicit none\ncontains\nsubroutine s()\ninteger :: x\nend subroutine s\nend module m\n"
);
c!(
    cf_host_assoc_26,
    "program t\ninteger :: x\nx=1\ncontains\nsubroutine s()\nprint *, x\nend subroutine s\nend program t\n"
);
c!(
    cf_spec_order_27,
    "program t\nimplicit none\ninteger :: x\nreal :: y\nend program t\n"
);
c!(
    cf_forward_ref_28,
    "program t\ninterface\nsubroutine s(x)\ninteger :: x\nend subroutine s\nend interface\nend program t\n"
);
c!(
    cf_block_data_29,
    "block data bd\ncommon /blk/ x\ninteger :: x\ndata x /1/\nend block data bd\n"
);
c!(
    cf_data_stmt_30,
    "program t\ninteger :: x\ndata x /1/\nend program t\n"
);
