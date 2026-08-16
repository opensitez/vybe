! vybe-test: fortran/generic_ambiguity/generic_ambiguity_10
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
integer :: hits = 0
interface g
module procedure s1,s2
end interface
contains
subroutine s1(c)
character(len=*)::c
hits = hits + 1
end
subroutine s2(l)
logical::l
end
end module m
program driver
use m
call g('a')
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program driver
