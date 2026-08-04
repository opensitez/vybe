! vybe-test: fortran/subroutine_extended/recursive_sum_one_to_five
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((series_sum(5)) /= 15) then
    print *, "FAIL: want [15] got [", series_sum(5), "]"
    stop 1
end if
contains
recursive function series_sum(n) result(s)
integer, intent(in) :: n
integer :: s
if (n <= 0) then
s = 0
else
s = n + series_sum(n - 1)
end if
end function series_sum
end program t
