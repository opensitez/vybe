! vybe-test: fortran/specification_part/spec_use_21
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
integer :: x
end module m
program p
use m
print *, x
end program p
