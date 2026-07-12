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
    spec_implicit_none_01,
    "program p
implicit none
integer :: x
x = 1
print *, x
end program p"
);
c!(
    spec_parameter_02,
    "program p
implicit none
integer, parameter :: n = 4
print *, n
end program p"
);
c!(
    spec_dimension_03,
    "program p
implicit none
integer, dimension(3) :: a
print *, 1
end program p"
);
c!(
    spec_save_04,
    "program p
implicit none
integer, save :: x
x = 1
print *, x
end program p"
);
c!(
    spec_data_05,
    "program p
implicit none
integer :: x
data x /1/
print *, x
end program p"
);
c!(
    spec_common_06,
    "program p
implicit none
integer :: x
common /blk/ x
print *, x
end program p"
);
c!(
    spec_equivalence_07,
    "program p
implicit none
integer :: a, b
equivalence (a, b)
print *, 1
end program p"
);
c!(
    spec_external_08,
    "program p
implicit none
external f
print *, 1
end program p"
);
c!(
    spec_intrinsic_09,
    "program p
implicit none
intrinsic abs
print *, abs(-1)
end program p"
);
c!(
    spec_pointer_10,
    "program p
implicit none
integer, pointer :: p
print *, 1
end program p"
);
c!(
    spec_target_11,
    "program p
implicit none
integer, target :: x
print *, 1
end program p"
);
c!(
    spec_allocatable_12,
    "program p
implicit none
integer, allocatable :: a(:)
print *, 1
end program p"
);
c!(
    spec_optional_13,
    "subroutine s(x)
implicit none
integer, optional :: x
end subroutine s"
);
c!(
    spec_intent_14,
    "subroutine s(x)
implicit none
integer, intent(in) :: x
end subroutine s"
);
c!(
    spec_kind_15,
    "program p
implicit none
integer(kind=4) :: x
print *, 1
end program p"
);
c!(
    spec_len_16,
    "program p
implicit none
character(len=8) :: s
print *, s
end program p"
);
c!(
    spec_type_17,
    "program p
implicit none
type :: t
 integer :: x
end type t
print *, 1
end program p"
);
c!(
    spec_public_18,
    "module m
implicit none
public :: x
integer :: x
end module m"
);
c!(
    spec_private_19,
    "module m
implicit none
private
integer :: x
end module m"
);
c!(
    spec_interface_20,
    "module m
implicit none
interface
 subroutine s(x)
  integer :: x
 end subroutine s
end interface
end module m"
);
c!(
    spec_use_21,
    "module m
implicit none
integer :: x
end module m
program p
use m
print *, x
end program p"
);
c!(
    spec_import_22,
    "module m
implicit none
contains
 subroutine s()
  import
 end subroutine s
end module m"
);
c!(
    spec_value_23,
    "subroutine s(x)
implicit none
integer, value :: x
end subroutine s"
);
c!(
    spec_protected_24,
    "module m
implicit none
integer, protected :: x
end module m"
);
c!(
    spec_bindc_25,
    "subroutine s() bind(c)
implicit none
end subroutine s"
);
c!(
    spec_volatile_26,
    "program p
implicit none
integer, volatile :: x
print *, 1
end program p"
);
c!(
    spec_asynchronous_27,
    "program p
implicit none
integer, asynchronous :: x
print *, 1
end program p"
);
c!(
    spec_contiguous_28,
    "subroutine s(a)
implicit none
integer, contiguous :: a(:)
end subroutine s"
);
c!(
    spec_codimension_29,
    "program p
implicit none
integer :: x[*]
print *, 1
end program p"
);
c!(
    spec_namelist_30,
    "program p
implicit none
integer :: x
namelist /n1/ x
print *, 1
end program p"
);
