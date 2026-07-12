use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(dtype_int_kind4_01, "program p
implicit none
integer(kind=4) :: x
print *, 1
end program p
");
c!(dtype_int_kind8_02, "program p
implicit none
integer(kind=8) :: x
print *, 1
end program p
");
c!(dtype_real_kind4_03, "program p
implicit none
real(kind=4) :: x
print *, 1
end program p
");
c!(dtype_real_kind8_04, "program p
implicit none
real(kind=8) :: x
print *, 1
end program p
");
c!(dtype_complex_kind8_05, "program p
implicit none
complex(kind=8) :: z
print *, 1
end program p
");
c!(dtype_logical_06, "program p
implicit none
logical :: l
print *, 1
end program p
");
c!(dtype_char_len_07, "program p
implicit none
character(len=5) :: s
print *, 1
end program p
");
c!(dtype_char_deferred_08, "program p
implicit none
character(len=:), allocatable :: s
print *, 1
end program p
");
c!(dtype_bool_expr_09, "program p
implicit none
logical :: l
l = .true.
print *, l
end program p
");
c!(dtype_int_array_10, "program p
implicit none
integer :: a(3)
print *, 1
end program p
");
c!(dtype_real_array_11, "program p
implicit none
real :: a(3)
print *, 1
end program p
");
c!(dtype_complex_array_12, "program p
implicit none
complex :: a(2)
print *, 1
end program p
");
c!(dtype_character_array_13, "program p
implicit none
character(len=4) :: a(2)
print *, 1
end program p
");
c!(dtype_derived_type_14, "program p
type :: t
 integer :: x
end type t
type(t) :: v
print *, 1
end program p
");
c!(dtype_extends_type_15, "type :: b
 integer :: x
end type b
type, extends(b) :: c
 integer :: y
end type c
");
c!(dtype_sequence_type_16, "type, sequence :: t
 integer :: x
end type t
");
c!(dtype_bindc_type_17, "type, bind(c) :: t
 integer :: x
end type t
");
c!(dtype_recursive_type_18, "type :: node
 integer :: x
 type(node), pointer :: next
end type node
");
c!(dtype_polymorphic_19, "program p
type :: t
 integer :: x
end type t
class(t), allocatable :: obj
print *, 1
end program p
");
c!(dtype_class_star_20, "subroutine s(x)
class(*) :: x
end subroutine s
");
c!(dtype_same_type_as_21, "program p
type :: t
 integer :: x
end type t
type(t) :: a, b
print *, same_type_as(a,b)
end program p
");
c!(dtype_extends_type_of_22, "program p
type :: t
 integer :: x
end type t
type(t) :: a, b
print *, extends_type_of(a,b)
end program p
");
c!(dtype_procedure_pointer_23, "program p
procedure(), pointer :: p
end program p
");
c!(dtype_c_ptr_24, "program p
use iso_c_binding
implicit none
type(c_ptr) :: p
print *, c_associated(p)
end program p
");
c!(dtype_c_funptr_25, "program p
use iso_c_binding
implicit none
type(c_funptr) :: fp
print *, 1
end program p
");
c!(dtype_alloc_comp_26, "type :: t
 integer, allocatable :: a(:)
end type t
");
c!(dtype_ptr_comp_27, "type :: t
 integer, pointer :: p
end type t
");
c!(dtype_codim_28, "program p
implicit none
integer :: x[*]
print *, 1
end program p
");
c!(dtype_assumed_type_29, "subroutine s(x)
type(*) :: x
end subroutine s
");
c!(dtype_assumed_rank_30, "subroutine s(a)
integer :: a(..)
end subroutine s
");