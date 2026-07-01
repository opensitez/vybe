use super::helpers::compile_ok;
macro_rules! c { ($name:ident,$src:expr)=>{ #[test] fn $name(){ compile_ok($src); } }; }
c!(if_explicit_01,"interface
subroutine s(x)
integer :: x
end subroutine s
end interface
");
c!(if_implicit_02,"external s
call s()
end
");
c!(if_block_03,"module m
interface
subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
");
c!(if_generic_04,"module m
interface g
module procedure s1
end interface
contains
subroutine s1()
end subroutine s1
end module m
");
c!(if_generic_two_05,"module m
interface g
module procedure si,sr
end interface
contains
subroutine si(i)
integer::i
end
subroutine sr(r)
real::r
end
end module m
");
c!(if_optional_06,"subroutine s(x)
integer, optional :: x
end subroutine s
");
c!(if_keyword_07,"subroutine s(x,y)
integer::x,y
end
program p
call s(y=2,x=1)
end program p
");
c!(if_positional_08,"subroutine s(x,y)
integer::x,y
end
program p
call s(1,2)
end program p
");
c!(if_altret_09,"subroutine s(*,*)
return 1
end
");
c!(if_result_10,"integer function f() result(r)
r=1
end function f
");
c!(if_recursive_result_11,"recursive integer function f(n) result(r)
integer::n
r=1
end function f
");
c!(if_proc_result_12,"function f() result(r)
integer :: r
r=1
end function f
");
c!(if_pass_13,"module m
type::t
contains
procedure,pass::s
end type
contains
subroutine s(this)
class(t)::this
end
end module m
");
c!(if_nopass_14,"module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s()
end
end module m
");
c!(if_intent_15,"subroutine s(x)
integer,intent(in)::x
end subroutine s
");
c!(if_value_16,"subroutine s(x)
integer,value::x
end subroutine s
");
c!(if_contiguous_17,"subroutine s(a)
real,contiguous::a(:)
end subroutine s
");
c!(if_target_18,"subroutine s(x)
integer,target::x
end subroutine s
");
c!(if_pointer_19,"subroutine s(x)
integer,pointer::x
end subroutine s
");
c!(if_allocatable_20,"subroutine s(x)
integer,allocatable::x(:)
end subroutine s
");
c!(if_abstract_21,"module m
abstract interface
subroutine s(x)
integer::x
end
end interface
end module m
");
c!(if_operator_22,"module m
interface operator(+)
module procedure addi
end interface
contains
integer function addi(a,b)
integer::a,b
addi=a+b
end
end module m
");
c!(if_assignment_23,"module m
interface assignment(=)
module procedure asg
end interface
contains
subroutine asg(a,b)
integer::a,b
a=b
end
end module m
");
c!(if_defined_io_24,"module m
type::t
contains
procedure::wf
generic::write(formatted)=>wf
end type
contains
subroutine wf(dtv,unit,iotype,v_list,iostat,iomsg)
class(t),intent(in)::dtv
integer,intent(in)::unit
character(len=*),intent(in)::iotype
integer,intent(in)::v_list(:)
integer,intent(out)::iostat
character(len=*),intent(inout)::iomsg
iostat=0
end
end module m
");
c!(if_binding_label_25,"subroutine s() bind(c,name='s_c')
end subroutine s
");
c!(if_b_independent_26,"subroutine s() bind(c)
end subroutine s
");
c!(if_dummy_proc_27,"subroutine apply(f)
external f
call f()
end subroutine apply
");
c!(if_proc_var_28,"program p
procedure(integer)::fp
end program p
");
c!(if_proc_ptr_29,"program p
procedure(),pointer::fp
end program p
");
c!(if_generic_resolution_30,"module m
interface g
module procedure s1,s2
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(r)
real::r
end
end module m
");
c!(if_explicit_fun_31,"interface
real function f(x)
real :: x
end function f
end interface
");
c!(if_module_sub_32,"module m
interface
module subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
");
c!(if_generic_three_33,"module m
interface g
module procedure s1,s2,s3
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(r)
real::r
end
subroutine s3(c)
complex::c
end
end module m
");
c!(if_operator_minus_34,"module m
interface operator(-)
module procedure subi
end interface
contains
integer function subi(a,b)
integer::a,b
subi=a-b
end
end module m
");
c!(if_assignment_real_int_35,"module m
interface assignment(=)
module procedure asgr
end interface
contains
subroutine asgr(a,b)
real::a
integer::b
a=real(b)
end
end module m
");
c!(if_pass_name_36,"module m
type::t
contains
procedure,pass(self)::s
end type
contains
subroutine s(self)
class(t)::self
end
end module m
");
c!(if_optional_two_37,"subroutine s(x,y)
integer, optional :: x,y
end subroutine s
");
c!(if_keyword_three_38,"subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(z=3,x=1,y=2)
end program p
");
c!(if_value_real_39,"subroutine s(x)
real, value :: x
end subroutine s
");
c!(if_pointer_proc_arg_40,"subroutine apply(f)
procedure() :: f
call f()
end subroutine apply
");