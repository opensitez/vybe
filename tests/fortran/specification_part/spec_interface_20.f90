! vybe-test: fortran/specification_part/spec_interface_20
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
interface
 subroutine s(x)
  integer :: x
 end subroutine s
end interface
end module m
