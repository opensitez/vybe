! vybe-test: fortran/subroutines/recursive_factorial
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program t
if ((fact(5)) /= 120) then
    print *, "FAIL: want [120] got [", fact(5), "]"
    stop 1
end if
contains
recursive function fact(n) result(r)
integer, intent(in) :: n
integer :: r
if (n <= 1) then
r = 1
else
r = n * fact(n - 1)
end if
end function fact
end program t
