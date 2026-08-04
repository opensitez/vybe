! vybe-test: fortran/subroutine_extended/recursive_power_two_to_fifth
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((ipow(2, 5)) /= 32) then
    print *, "FAIL: want [32] got [", ipow(2, 5), "]"
    stop 1
end if
contains
recursive function ipow(base, exp) result(r)
integer, intent(in) :: base, exp
integer :: r
if (exp == 0) then
r = 1
else
r = base * ipow(base, exp - 1)
end if
end function ipow
end program t
