! vybe-test: fortran/generic_resolution/generic_resolution_05
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
module m
interface g
module procedure si
end interface
contains
subroutine si(i)
integer::i
end
end module m
program p
use m
call g(1)
end program p
