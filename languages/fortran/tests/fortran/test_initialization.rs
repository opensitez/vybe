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
    init_default_01,
    "program p
integer :: x=1
print *,x
end program p
"
);
c!(
    init_component_02,
    "type::t
integer::x=1
end type t
program p
type(t)::v
print *,v%x
end program p
"
);
c!(
    init_parameter_03,
    "program p
integer,parameter::x=1
print *,x
end program p
"
);
c!(
    init_save_04,
    "program p
integer,save::x=1
print *,x
end program p
"
);
c!(
    init_data_05,
    "program p
integer::x
data x/1/
print *,x
end program p
"
);
c!(
    init_block_data_06,
    "block data bd
integer::x
common /blk/ x
data x/1/
end block data bd
"
);
c!(
    init_array_07,
    "program p
integer::a(3)=[1,2,3]
print *,a
end program p
"
);
c!(
    init_char_08,
    "program p
character(len=3)::s='abc'
print *,s
end program p
"
);
c!(
    init_logical_09,
    "program p
logical::l=.true.
print *,l
end program p
"
);
c!(
    init_complex_10,
    "program p
complex::z=(1.0,2.0)
print *,z
end program p
"
);
c!(
    init_derived_11,
    "type::t
integer::x=1
end type t
program p
type(t)::v=t(2)
print *,v%x
end program p
"
);
c!(
    init_pointer_null_12,
    "program p
integer,pointer::p=>null()
end program p
"
);
c!(
    init_alloc_char_13,
    "program p
character(len=:),allocatable::s
allocate(character(len=3)::s)
s='abc'
end program p
"
);
c!(
    init_common_14,
    "program p
integer::x
common /blk/ x
x=1
end program p
"
);
c!(
    init_equivalence_15,
    "program p
integer::a,b
equivalence(a,b)
a=1
print *,b
end program p
"
);
c!(
    init_structure_ctor_16,
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
    init_nested_17,
    "type::u
integer::y=2
end type u
type::t
type(u)::u1
end type t
program p
type(t)::v
end program p
"
);
c!(
    init_real_18,
    "program p
real::x=1.5
print *,x
end program p
"
);
c!(
    init_kind_19,
    "program p
integer(kind=8)::x=1_8
print *,x
end program p
"
);
c!(
    init_array_data_20,
    "program p
integer::a(3)
data a/1,2,3/
print *,a
end program p
"
);

c!(
    init_component_array_21,
    "type::t
integer :: a(3) = [1,2,3]
end type t
program p
type(t)::v
print *, v%a(2)
end program p
"
);

c!(
    init_component_char_22,
    "type::t
character(len=4) :: tag = 'init'
logical :: active = .true.
end type t
program p
type(t)::v
print *, v%tag
print *, v%active
end program p
"
);

c!(
    init_parameter_expression_23,
    "program p
integer, parameter :: one = 1
integer, parameter :: two = one + 1
integer, parameter :: three = two + one
print *, three
end program p
"
);

c!(
    init_array_with_implied_shape_24,
    "program p
integer, dimension(2,3) :: m = reshape([1,2,3,4,5,6], [2,3])
print *, m(2,2)
end program p
"
);

c!(
    init_pointer_target_default_25,
    "program p
integer, target :: t = 7
integer, pointer :: p => t
print *, p
end program p
"
);

c!(
    init_allocatable_default_26,
    "program p
integer, allocatable :: a(:)
a = [1,2,3]
print *, a(1)
end program p
"
);

c!(
    init_derived_optional_component_27,
    "type::inner
integer :: x = 10
end type inner
type::outer
type(inner) :: c
integer :: y = 4
end type outer
program p
type(outer)::v
print *, v%c%x
print *, v%y
end program p
"
);
