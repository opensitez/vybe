! vybe-test: fortran/subroutine_extended/recursive_digit_sum_four_digits
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((digit_sum(9876)) /= 30) then
    print *, "FAIL: want [30] got [", digit_sum(9876), "]"
    stop 1
end if
contains
recursive function digit_sum(n) result(s)
integer, intent(in) :: n
integer :: s
if (n < 10) then
s = n
else
s = mod(n, 10) + digit_sum(n / 10)
end if
end function digit_sum
end program t
