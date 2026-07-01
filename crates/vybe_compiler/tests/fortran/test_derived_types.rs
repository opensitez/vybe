use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(dt_type_ext_01,"type::b
integer::x
end type b
type,extends(b)::c
integer::y
end type c
");
c!(dt_final_02,"type::t
integer::x
contains
final::fin
end type t
contains
subroutine fin(x)
type(t)::x
end subroutine fin
");
c!(dt_bound_generic_03,"module m
type::t
contains
generic::g=>s
procedure::s
end type t
contains
subroutine s(this)
class(t)::this
end subroutine s
end module m
");
c!(dt_bound_op_04,"module m
type::t
contains
procedure::add
generic::operator(+)=>add
end type t
contains
integer function add(this,other)
class(t)::this
class(t)::other
add=0
end function add
end module m
");
c!(dt_private_comp_05,"module m
type::t
private
integer::x
end type t
end module m
");
c!(dt_public_comp_06,"module m
type::t
public
integer::x
end type t
end module m
");
c!(dt_sequence_07,"type,sequence::t
integer::x
end type t
");
c!(dt_bindc_08,"type,bind(c)::t
integer::x
end type t
");
c!(dt_recursive_09,"type::node
integer::x
type(node),pointer::next
end type node
");
c!(dt_poly_comp_10,"type::t
class(*), allocatable :: item
end type t
");
c!(dt_class_default_11,"program p
type::t
integer::x
end type t
class(t),allocatable::o
allocate(o)
end program p
");
c!(dt_class_star_12,"subroutine s(x)
class(*)::x
end subroutine s
");
c!(dt_same_type_13,"program p
type::t
integer::x
end type t
type(t)::a,b
print *, same_type_as(a,b)
end program p
");
c!(dt_extends_type_14,"program p
type::t
integer::x
end type t
type(t)::a,b
print *, extends_type_of(a,b)
end program p
");
c!(dt_select_type_15,"program p
class(*), allocatable :: x
allocate(integer :: x)
select type(x)
 type is(integer)
  print *, x
 class default
end select
end program p
");
c!(dt_constructor_16,"type::t
integer::x
end type t
program p
type(t)::v
v=t(1)
print *,v%x
end program p
");
c!(dt_default_init_17,"type::t
integer::x=1
end type t
program p
type(t)::v
print *,v%x
end program p
");
c!(dt_comp_init_18,"type::t
integer::x=1
real::y=2.0
end type t
program p
type(t)::v
print *,v%x
end program p
");
c!(dt_type_array_19,"type::t
integer::x
end type t
program p
type(t)::a(2)
print *,1
end program p
");
c!(dt_nested_type_20,"type::t1
integer::x
end type t1
type::t2
type(t1)::a
end type t2
");