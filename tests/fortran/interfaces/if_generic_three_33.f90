! vybe-test: fortran/interfaces/if_generic_three_33
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
integer :: chosen = 0
interface g
module procedure s1,s2,s3
end interface
contains
subroutine s1(i)
integer::i
chosen = 1
end
subroutine s2(r)
real::r
chosen = 2
end
subroutine s3(c)
complex::c
chosen = 3
end
end module m
program t
use m
call g(1)
if (chosen /= 1) then
    print *, "FAIL: want [1] got [", chosen, "]"
    stop 1
end if
call g(1.0)
if (chosen /= 2) then
    print *, "FAIL: want [2] got [", chosen, "]"
    stop 1
end if
call g((1.0, 2.0))
if (chosen /= 3) then
    print *, "FAIL: want [3] got [", chosen, "]"
    stop 1
end if
end program t
