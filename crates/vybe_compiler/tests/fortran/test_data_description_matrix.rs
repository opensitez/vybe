use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(dd_integer_01, "program p
implicit none
integer :: x
print *, 1
end program p
");
c!(dd_real_02, "program p
implicit none
real :: x
print *, 1
end program p
");
c!(dd_complex_03, "program p
implicit none
complex :: z
print *, 1
end program p
");
c!(dd_logical_04, "program p
implicit none
logical :: l
print *, 1
end program p
");
c!(dd_character_05, "program p
implicit none
character(len=10) :: s
print *, 1
end program p
");
c!(dd_parameter_06, "program p
implicit none
integer, parameter :: n = 4
print *, n
end program p
");
c!(dd_dimension_07, "program p
implicit none
integer, dimension(3) :: a
print *, 1
end program p
");
c!(dd_allocatable_08, "program p
implicit none
integer, allocatable :: a(:)
print *, 1
end program p
");
c!(dd_pointer_09, "program p
implicit none
integer, pointer :: p
print *, 1
end program p
");
c!(dd_target_10, "program p
implicit none
integer, target :: x
print *, 1
end program p
");
c!(dd_save_11, "program p
implicit none
integer, save :: x
print *, 1
end program p
");
c!(dd_protected_12, "module m
implicit none
integer, protected :: x
end module m
");
c!(dd_volatile_13, "program p
implicit none
integer, volatile :: x
print *, 1
end program p
");
c!(dd_async_14, "program p
implicit none
integer, asynchronous :: x
print *, 1
end program p
");
c!(dd_bindc_15, "type, bind(c) :: t
 integer :: x
end type t
");
c!(dd_sequence_16, "type, sequence :: t
 integer :: x
end type t
");
c!(dd_extends_17, "type :: base
 integer :: x
end type base
type, extends(base) :: child
 integer :: y
end type child
");
c!(dd_abstract_18, "type, abstract :: t
 integer :: x
end type t
");
c!(dd_deferred_19, "type, abstract :: t
contains
 procedure(p), deferred :: s
end type t
abstract interface
 subroutine p(this)
  import :: t
  class(t) :: this
 end subroutine p
end interface
");
c!(dd_private_public_20, "module m
implicit none
private
public :: x
integer :: x
end module m
");
c!(dd_optional_21, "subroutine s(x)
integer, optional :: x
end subroutine s
");
c!(dd_value_22, "subroutine s(x)
integer, value :: x
end subroutine s
");
c!(dd_codim_23, "program p
implicit none
integer :: x[*]
print *, 1
end program p
");
c!(dd_deferred_shape_24, "subroutine s(a)
integer, allocatable :: a(:)
end subroutine s
");
c!(dd_assumed_shape_25, "subroutine s(a)
integer :: a(:)
end subroutine s
");
c!(dd_assumed_rank_26, "subroutine s(a)
integer :: a(..)
end subroutine s
");
c!(dd_assumed_type_27, "subroutine s(x)
type(*) :: x
end subroutine s
");
c!(dd_explicit_shape_28, "subroutine s(a)
integer :: a(3)
end subroutine s
");
c!(dd_deferred_len_char_29, "program p
implicit none
character(len=:), allocatable :: s
print *, 1
end program p
");
c!(dd_proc_pointer_30, "program p
procedure(), pointer :: fp
end program p
");