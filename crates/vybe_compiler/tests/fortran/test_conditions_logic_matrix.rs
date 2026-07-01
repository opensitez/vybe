use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(cond_if_01, "program t\ninteger :: x=1\nif (x==1) print *, x\nend program t\n");
c!(cond_if_else_02, "program t\ninteger :: x=0\nif (x==1) then\n print *, 1\nelse\n print *, 0\nend if\nend program t\n");
c!(cond_nested_03, "program t\ninteger :: a=1,b=2\nif (a==1) then\n if (b==2) print *, 1\nend if\nend program t\n");
c!(cond_and_04, "program t\nlogical :: x\nx = (1<2 .and. 2<3)\nprint *, x\nend program t\n");
c!(cond_or_05, "program t\nlogical :: x\nx = (1>2 .or. 2<3)\nprint *, x\nend program t\n");
c!(cond_not_06, "program t\nlogical :: x\nx = .not.(1>2)\nprint *, x\nend program t\n");
c!(cond_eqv_07, "program t\nlogical :: x\nx = .true. .eqv. .true.\nprint *, x\nend program t\n");
c!(cond_neqv_08, "program t\nlogical :: x\nx = .true. .neqv. .false.\nprint *, x\nend program t\n");
c!(cond_select_case_09, "program t\ninteger :: x=2\nselect case (x)\ncase(1)\n print *, 1\ncase(2)\n print *, 2\nend select\nend program t\n");
c!(cond_case_range_10, "program t\ninteger :: x=5\nselect case (x)\ncase(1:10)\n print *, x\nend select\nend program t\n");
c!(cond_case_default_11, "program t\ninteger :: x=20\nselect case (x)\ncase default\n print *, x\nend select\nend program t\n");
c!(cond_arith_if_12, "program t\ninteger :: x=1\nif (x) 10,20,30\n10 continue\n20 continue\n30 continue\nend program t\n");
c!(cond_associated_13, "program t\ninteger,target :: x\ninteger,pointer :: p\np => x\nif (associated(p)) print *, 1\nend program t\n");
c!(cond_allocated_14, "program t\ninteger, allocatable :: a(:)\nallocate(a(2))\nif (allocated(a)) print *, 1\nend program t\n");
c!(cond_present_15, "subroutine s(x)\ninteger, optional :: x\nif (present(x)) print *, x\nend subroutine s\n");
c!(cond_same_type_16, "type :: t\n integer :: x\nend type t\nprogram p\nclass(*), allocatable :: a\nallocate(t::a)\nselect type(a)\ntype is (t)\n print *, 1\nend select\nend program p\n");
c!(cond_class_default_17, "type :: t\n integer :: x\nend type t\nprogram p\nclass(*), allocatable :: a\nallocate(t::a)\nselect type(a)\nclass default\n print *, 1\nend select\nend program p\n");
c!(cond_where_18, "program t\ninteger :: a(3)=[1,2,3]\nwhere (a>1) a=0\nend program t\n");
c!(cond_elsewhere_19, "program t\ninteger :: a(3)=[1,2,3]\nwhere (a>1)\n a=0\nelsewhere\n a=1\nend where\nend program t\n");
c!(cond_merge_20, "program t\ninteger :: x\nx = merge(1,2,.true.)\nend program t\n");
c!(cond_char_compare_21, "program t\ncharacter(len=3) :: s='abc'\nif (s=='abc') print *, 1\nend program t\n");
c!(cond_real_compare_22, "program t\nreal :: x=1.0\nif (x>=1.0) print *, 1\nend program t\n");
c!(cond_complex_eq_23, "program t\ncomplex :: z=(1.0,2.0)\nif (z==(1.0,2.0)) print *, 1\nend program t\n");
c!(cond_kind_expr_24, "program t\ninteger(kind=8) :: x=1_8\nif (x==1_8) print *, 1\nend program t\n");
c!(cond_optional_else_25, "subroutine s(x)\ninteger, optional :: x\nif (.not.present(x)) print *, 0\nend subroutine s\n");
c!(cond_do_while_26, "program t\ninteger :: i=0\ndo while (i<2)\n i=i+1\nend do\nend program t\n");
c!(cond_exit_cycle_27, "program t\ninteger :: i\ndo i=1,3\n if (i==2) cycle\n if (i==3) exit\nend do\nend program t\n");
c!(cond_block_if_28, "program t\nlogical :: ok=.true.\nif (ok) then\n print *, 1\nend if\nend program t\n");
c!(cond_mask_array_29, "program t\nlogical :: m(3)\nm = [1,2,3] > 1\nprint *, m\nend program t\n");
c!(cond_any_all_30, "program t\nlogical :: m(3)=[.true.,.false.,.true.]\nprint *, any(m), all(m)\nend program t\n")