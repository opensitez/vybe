! vybe-test: fortran/interfaces/if_generic_two_05
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
integer :: chosen = 0
interface g
module procedure si,sr
end interface
contains
subroutine si(i)
integer::i
chosen = 1
end
subroutine sr(r)
real::r
chosen = 2
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
end program t
