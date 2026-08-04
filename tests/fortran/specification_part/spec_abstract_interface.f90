! vybe-test: fortran/specification_part/spec_abstract_interface
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
abstract interface
 integer function f(x)
  integer, intent(in) :: x
 end function f
end interface
end module m
