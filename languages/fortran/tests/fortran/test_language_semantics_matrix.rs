use super::helpers::compile_ok;

macro_rules! c {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

c!(
    sem_scope_01,
    "program t\ninteger :: x\nx=1\ncontains\nsubroutine s()\nprint *, x\nend subroutine s\nend program t\n"
);
c!(
    sem_block_scope_02,
    "program t\nblock\ninteger :: x\nx=1\nprint *, x\nend block\nend program t\n"
);
c!(
    sem_associate_scope_03,
    "program t\ninteger :: x=1\nassociate (a => x)\nprint *, a\nend associate\nend program t\n"
);
c!(
    sem_module_scope_04,
    "module m\ninteger :: x\nend module m\nprogram t\nuse m\nx=1\nend program t\n"
);
c!(
    sem_private_vis_05,
    "module m\nprivate\ninteger :: x\nend module m\n"
);
c!(
    sem_public_vis_06,
    "module m\npublic\ninteger :: x\nend module m\n"
);
c!(
    sem_rename_use_07,
    "module m\ninteger :: x\nend module m\nprogram t\nuse m, y => x\ny=1\nend program t\n"
);
c!(
    sem_only_use_08,
    "module m\ninteger :: x\ninteger :: z\nend module m\nprogram t\nuse m, only: x\nx=1\nend program t\n"
);
c!(
    sem_host_assoc_09,
    "program t\ninteger :: x=1\ncontains\nsubroutine s()\nprint *, x\nend subroutine s\nend program t\n"
);
c!(
    sem_name_resolution_10,
    "program t\ninteger :: x\nreal :: y\nx=1\ny=2.0\nend program t\n"
);
c!(
    sem_type_compat_11,
    "program t\ninteger :: x\nreal :: y\nx=1\ny=x\nend program t\n"
);
c!(
    sem_assign_rules_12,
    "program t\ninteger :: x\nreal :: y\ny=1.0\nx=int(y)\nend program t\n"
);
c!(
    sem_eval_order_13,
    "program t\ninteger :: a=1,b=2,c\nc = a + b\nend program t\n"
);
c!(
    sem_aliasing_14,
    "program t\ninteger, target :: x\ninteger, pointer :: p\np => x\np = 1\nend program t\n"
);
c!(
    sem_init_15,
    "program t\ninteger :: x=1\nprint *, x\nend program t\n"
);
c!(
    sem_conversion_16,
    "program t\ninteger :: x\nreal :: y=1.5\nx = int(y)\nend program t\n"
);
c!(
    sem_calling_17,
    "subroutine s(x)\ninteger :: x\nend subroutine s\nprogram t\ninteger :: x=1\ncall s(x)\nend program t\n"
);
c!(
    sem_lifetime_18,
    "program t\ncontains\nsubroutine s()\ninteger, save :: x=1\nprint *, x\nend subroutine s\nend program t\n"
);
c!(
    sem_storage_19,
    "program t\ninteger, save :: x\nx=1\nend program t\n"
);
c!(
    sem_common_20,
    "program t\ncommon /blk/ x\ninteger :: x\nx=1\nend program t\n"
);
c!(
    sem_equivalence_21,
    "program t\ninteger :: a\nreal :: b\nequivalence (a,b)\nend program t\n"
);
c!(
    sem_construct_scope_22,
    "program t\ninteger :: x=1\nif (x==1) then\ninteger :: y\ny=2\nend if\nend program t\n"
);
c!(
    sem_module_host_23,
    "module m\ninteger :: x=1\ncontains\nsubroutine s()\nprint *, x\nend subroutine s\nend module m\n"
);
c!(
    sem_pointer_target_24,
    "program t\ninteger, target :: x\ninteger, pointer :: p\np => x\nend program t\n"
);
c!(
    sem_proc_pointer_25,
    "program t\nprocedure(), pointer :: p\nend program t\n"
);
c!(
    sem_object_life_26,
    "type :: t\ninteger :: x\nend type t\nprogram p\ntype(t) :: v\nv%x=1\nend program p\n"
);
c!(
    sem_default_init_27,
    "type :: t\ninteger :: x=1\nend type t\nprogram p\ntype(t) :: v\nprint *, v%x\nend program p\n"
);
c!(
    sem_component_init_28,
    "type :: t\ninteger :: x=1\nreal :: y=2.0\nend type t\nprogram p\ntype(t) :: v\nend program p\n"
);
c!(
    sem_parameter_init_29,
    "program t\ninteger, parameter :: x=1\nprint *, x\nend program t\n"
);
c!(
    sem_standard_conf_30,
    "program t\nimplicit none\ninteger :: x\nx=1\nend program t\n"
);
