! vybe-test: fortran/generic_interfaces/gen_if_05
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
integer :: hits = 0
interface g
module procedure s1,s2,s3
end interface
contains
subroutine s1(i)
integer::i
hits = hits + 1
end
subroutine s2(r)
real::r
end
subroutine s3(c)
complex::c
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
