use super::helpers::compile_ok;

macro_rules! c {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() { compile_ok($src); }
    };
}

c!(program_basic_01, "program p
print *, 1
end program p
");
c!(program_contains_02, "program p
contains
subroutine s()
print *, 1
end subroutine s
end program p
");
c!(program_internal_proc_03, "program p
call s()
contains
subroutine s()
print *, 1
end subroutine s
end program p
");
c!(program_recursive_04, "recursive subroutine s()
print *, 1
end subroutine s
");
c!(program_function_05, "integer function f()
f = 1
end function f
");
c!(program_module_proc_06, "module m
contains
subroutine s()
print *, 1
end subroutine s
end module m
");
c!(program_submodule_style_07, "module m
interface
module subroutine s()
end subroutine s
end interface
end module m
");
c!(program_statement_fn_08, "program p
implicit none
integer :: f
f(x) = x + 1
print *, f(1)
end program p
");
c!(program_abstract_interface_09, "module m
abstract interface
subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
");
c!(program_generic_interface_10, "module m
interface g
module procedure s1
end interface
contains
subroutine s1()
print *, 1
end subroutine s1
end module m
");
c!(program_operator_interface_11, "module m
interface operator(+)
module procedure addi
end interface
contains
integer function addi(a,b)
integer :: a,b
addi = a+b
end function addi
end module m
");
c!(program_assignment_interface_12, "module m
interface assignment(=)
module procedure assigni
end interface
contains
subroutine assigni(a,b)
integer :: a,b
a = b
end subroutine assigni
end module m
");
c!(program_defined_io_13, "module m
type :: t
 integer :: x
contains
 procedure :: write_formatted
 generic :: write(formatted) => write_formatted
end type t
contains
subroutine write_formatted(dtv, unit, iotype, v_list, iostat, iomsg)
 class(t), intent(in) :: dtv
 integer, intent(in) :: unit
 character(len=*), intent(in) :: iotype
 integer, intent(in) :: v_list(:)
 integer, intent(out) :: iostat
 character(len=*), intent(inout) :: iomsg
 iostat = 0
end subroutine write_formatted
end module m
");
c!(program_binding_label_14, "subroutine s() bind(c, name='s_c')
end subroutine s
");
c!(program_b_independent_15, "subroutine s() bind(c)
end subroutine s
");
c!(program_external_proc_16, "external s
call s()
end
");
c!(program_dummy_proc_17, "subroutine apply(f)
external f
call f()
end subroutine apply
");
c!(program_proc_pointer_18, "program p
procedure(), pointer :: fp
end program p
");
c!(program_proc_variable_19, "program p
procedure(integer) :: fp
end program p
");
c!(program_interface_result_20, "interface
integer function f()
end function f
end interface
");
c!(program_generic_resolution_21, "module m
interface g
module procedure si, sr
end interface
contains
subroutine si(i)
integer :: i
end subroutine si
subroutine sr(r)
real :: r
end subroutine sr
end module m
");
c!(program_optional_args_22, "subroutine s(x)
integer, optional :: x
end subroutine s
");
c!(program_keyword_args_23, "subroutine s(x,y)
integer :: x,y
end subroutine s
program p
call s(y=2, x=1)
end program p
");
c!(program_result_var_24, "integer function f() result(r)
r = 1
end function f
");
c!(program_recursive_result_25, "recursive integer function f(n) result(r)
integer :: n
if (n <= 1) then
 r = 1
else
 r = n * f(n-1)
end if
end function f
");
c!(program_pass_nopass_26, "module m
type :: t
contains
 procedure, pass :: s1
 procedure, nopass :: s2
end type t
contains
subroutine s1(this)
 class(t) :: this
end subroutine s1
subroutine s2()
end subroutine s2
end module m
");
c!(program_contiguous_arg_27, "subroutine s(a)
real, contiguous :: a(:)
end subroutine s
");
c!(program_target_arg_28, "subroutine s(a)
integer, target :: a
end subroutine s
");
c!(program_pointer_arg_29, "subroutine s(a)
integer, pointer :: a
end subroutine s
");
c!(program_allocatable_arg_30, "subroutine s(a)
integer, allocatable :: a(:)
end subroutine s
");
c!(program_main_with_module_31, "module m
integer :: x=1
end module m
program p
use m
print *, x
end program p
");
c!(program_main_with_block_32, "program p
block
 print *, 1
end block
end program p
");
c!(program_main_with_associate_33, "program p
integer :: x=1
associate(y=>x)
 print *, y
end associate
end program p
");
c!(program_module_function_34, "module m
contains
integer function f()
f=1
end function f
end module m
");
c!(program_module_subroutine_35, "module m
contains
subroutine s()
print *, 1
end subroutine s
end module m
");
c!(program_main_call_internal_36, "program p
call s()
contains
subroutine s()
 print *, 1
end subroutine s
end program p
");
c!(program_result_real_37, "real function f() result(r)
r = 1.0
end function f
");
c!(program_result_complex_38, "complex function f() result(r)
r = (1.0,2.0)
end function f
");
c!(program_contains_function_39, "program p
print *, f()
contains
integer function f()
f = 1
end function f
end program p
");
c!(program_dummy_external_40, "subroutine apply(f)
external f
call f()
end subroutine apply
");