use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(proc_sub_01, "subroutine s()\nprint *, 1\nend subroutine s\n");
c!(proc_func_02, "integer function f()\nf = 1\nend function f\n");
c!(proc_prog_call_03, "program t\ncall s()\ncontains\nsubroutine s()\nprint *, 1\nend subroutine s\nend program t\n");
c!(proc_internal_04, "program t\ncall s\ncontains\nsubroutine s\nprint *, 1\nend subroutine s\nend program t\n");
c!(proc_module_05, "module m\ncontains\nsubroutine s()\nprint *, 1\nend subroutine s\nend module m\n");
c!(proc_recursive_06, "recursive subroutine s(n)\ninteger :: n\nif (n>0) call s(n-1)\nend subroutine s\n");
c!(proc_result_07, "integer function f() result(r)\nr = 1\nend function f\n");
c!(proc_optional_08, "subroutine s(x)\ninteger, optional :: x\nif (present(x)) print *, x\nend subroutine s\n");
c!(proc_keyword_09, "subroutine s(x,y)\ninteger :: x,y\nend subroutine s\nprogram t\ncall s(y=2,x=1)\nend program t\n");
c!(proc_intent_10, "subroutine s(x)\ninteger, intent(inout) :: x\nx = x + 1\nend subroutine s\n");
c!(proc_value_11, "subroutine s(x)\ninteger, value :: x\nprint *, x\nend subroutine s\n");
c!(proc_pointer_arg_12, "subroutine s(p)\ninteger, pointer :: p\nend subroutine s\n");
c!(proc_target_arg_13, "subroutine s(x)\ninteger, target :: x\nend subroutine s\n");
c!(proc_alloc_arg_14, "subroutine s(a)\ninteger, allocatable :: a(:)\nend subroutine s\n");
c!(proc_contiguous_15, "subroutine s(a)\ninteger, contiguous :: a(:)\nend subroutine s\n");
c!(proc_bindc_16, "subroutine s() bind(c)\nend subroutine s\n");
c!(proc_interface_17, "interface\nsubroutine s(x)\ninteger :: x\nend subroutine s\nend interface\n");
c!(proc_abstract_18, "abstract interface\nsubroutine s(x)\ninteger :: x\nend subroutine s\nend interface\n");
c!(proc_generic_19, "module m\ninterface s\n module procedure s1\nend interface\ncontains\nsubroutine s1()\nend subroutine s1\nend module m\n");
c!(proc_operator_20, "module m\ninterface operator(.foo.)\n module procedure f\nend interface\ncontains\nlogical function f(a,b)\nlogical :: a,b\nf = a .or. b\nend function f\nend module m\n");
c!(proc_assignment_21, "module m\ninterface assignment(=)\n module procedure s\nend interface\ncontains\nsubroutine s(a,b)\ninteger :: a,b\na=b\nend subroutine s\nend module m\n");
c!(proc_dummy_22, "subroutine s(p)\nexternal :: p\ncall p()\nend subroutine s\n");
c!(proc_proc_ptr_23, "program t\nprocedure(), pointer :: p\nend program t\n");
c!(proc_elem_24, "elemental integer function f(x)\ninteger, intent(in) :: x\nf=x+1\nend function f\n");
c!(proc_pure_25, "pure integer function f(x)\ninteger, intent(in) :: x\nf=x+1\nend function f\n");
c!(proc_impure_26, "impure elemental integer function f(x)\ninteger, intent(in) :: x\nf=x+1\nend function f\n");
c!(proc_module_call_27, "module m\ncontains\nsubroutine s()\nprint *, 1\nend subroutine s\nend module m\nprogram t\nuse m\ncall s()\nend program t\n");
c!(proc_result_char_28, "character(len=3) function f()\nf = 'abc'\nend function f\n");
c!(proc_recursive_func_29, "recursive integer function f(n) result(r)\ninteger :: n\nif (n<=0) then\n r=0\nelse\n r=f(n-1)\nend if\nend function f\n");
c!(proc_contains_chain_30, "program t\ncall a()\ncontains\nsubroutine a()\ncall b()\nend subroutine a\nsubroutine b()\nprint *, 1\nend subroutine b\nend program t\n")