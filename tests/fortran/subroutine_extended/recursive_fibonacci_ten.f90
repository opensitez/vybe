! vybe-test: fortran/subroutine_extended/recursive_fibonacci_ten
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((fib(10)) /= 55) then
    print *, "FAIL: want [55] got [", fib(10), "]"
    stop 1
end if
contains
recursive function fib(n) result(r)
integer, intent(in) :: n
integer :: r
if (n <= 1) then
r = n
else
r = fib(n - 1) + fib(n - 2)
end if
end function fib
end program t
