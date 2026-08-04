! vybe-test: fortran/subroutine_extended/intent_inout_halve_integer
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: x
x = 20
call halve(x)
if ((x) /= 10) then
    print *, "FAIL: want [10] got [", x, "]"
    stop 1
end if
contains
subroutine halve(n)
integer, intent(inout) :: n
n = n / 2
end subroutine halve
end program t
