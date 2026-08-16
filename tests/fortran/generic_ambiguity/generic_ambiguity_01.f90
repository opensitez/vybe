! vybe-test: fortran/generic_ambiguity/generic_ambiguity_01
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
integer :: hits = 0
interface g
module procedure si,sr
end interface
contains
subroutine si(i)
integer::i
hits = hits + 1
end
subroutine sr(r)
real::r
end
end module m
program driver
use m
call g(3)
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
