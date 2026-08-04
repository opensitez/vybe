! vybe-test: fortran/subroutine_extended/recursive_factorial_four
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((fact4(4)) /= 24) then
    print *, "FAIL: want [24] got [", fact4(4), "]"
    stop 1
end if
contains
recursive function fact4(n) result(r)
integer, intent(in) :: n
integer :: r
if (n <= 1) then
r = 1
else
r = n * fact4(n - 1)
end if
end function fact4
end program t
