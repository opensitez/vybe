use super::helpers::compile_ok;
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    oop_dispatch_01,
    "module m
type::b
contains
procedure::show
end type b
contains
subroutine show(this)
class(b)::this
end subroutine show
end module m
"
);
c!(
    oop_override_02,
    "module m
type::b
contains
procedure::show
end type b
type,extends(b)::c
contains
procedure::show=>show_c
end type c
contains
subroutine show(this)
class(b)::this
end
subroutine show_c(this)
class(c)::this
end
end module m
"
);
c!(
    oop_inherit_chain_03,
    "type::a
integer::x
end type a
type,extends(a)::b
integer::y
end type b
type,extends(b)::c
integer::z
end type c
"
);
c!(
    oop_factory_04,
    "module m
type::t
integer::x
end type t
contains
function make() result(r)
type(t)::r
r%x=1
end function make
end module m
"
);
c!(
    oop_constructor_05,
    "type::t
integer::x
end type t
program p
type(t)::v
v=t(1)
end program p
"
);
c!(
    oop_class_assign_06,
    "type::t
integer::x
end type t
program p
class(t),allocatable::a,b
allocate(a,b)
a=b
end program p
"
);
c!(
    oop_class_arg_07,
    "subroutine s(x)
class(*), intent(in) :: x
end subroutine s
"
);
c!(
    oop_class_result_08,
    "function f() result(r)
class(*), allocatable :: r
allocate(integer :: r)
end function f
"
);
c!(
    oop_class_array_09,
    "type::t
integer::x
end type t
program p
class(t), allocatable :: a(:)
allocate(a(2))
end program p
"
);
c!(
    oop_lifetime_10,
    "type::t
integer::x
end type t
program p
block
type(t)::v
v%x=1
end block
end program p
"
);
c!(
    oop_self_ref_11,
    "module m
type::t
contains
procedure::show
end type t
contains
subroutine show(this)
class(t)::this
print *,1
end subroutine show
end module m
"
);
c!(
    oop_super_ref_12,
    "module m
type::b
contains
procedure::show
end type b
type,extends(b)::c
contains
procedure::show=>show_c
end type c
contains
subroutine show(this)
class(b)::this
end
subroutine show_c(this)
class(c)::this
end
end module m
"
);
c!(
    oop_property_like_13,
    "type::t
integer::x
contains
procedure::getx
end type t
contains
integer function getx(this)
class(t)::this
getx=this%x
end function getx
"
);
c!(
    oop_encap_14,
    "module m
type::t
private
integer::x
contains
procedure::setx
end type t
contains
subroutine setx(this,v)
class(t)::this
integer::v
this%x=v
end
end module m
"
);
c!(
    oop_abstract_15,
    "type,abstract::t
contains
procedure(p),deferred::run
end type t
abstract interface
subroutine p(this)
import t
class(t)::this
end
end interface
"
);
c!(
    oop_final_16,
    "type::t
contains
final::fin
end type t
contains
subroutine fin(x)
type(t)::x
end subroutine fin
"
);
c!(
    oop_bound_generic_17,
    "module m
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
"
);
c!(
    oop_bound_op_18,
    "module m
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
"
);
c!(
    oop_non_over_19,
    "module m
type::t
contains
procedure,non_overridable::s
end type t
contains
subroutine s(this)
class(t)::this
end subroutine s
end module m
"
);
c!(
    oop_polymorphic_comp_20,
    "type::box
class(*), allocatable :: item
end type box
"
);
