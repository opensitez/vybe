! vybe-test: fortran/generic_resolution/generic_resolution_06
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
module m
interface g
module procedure sr
end interface
contains
subroutine sr(r)
real::r
end
end module m
program p
use m
call g(1.0)
end program p
