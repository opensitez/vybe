! vybe-test: fortran/subroutine_extended/internal_sub_calls_pure_function
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call show_diff(9, 4)
contains
pure function pdiff(a, b) result(d)
integer, intent(in) :: a, b
integer :: d
d = a - b
end function pdiff
subroutine show_diff(x, y)
integer, intent(in) :: x, y
if ((pdiff(x, y)) /= 5) then
    print *, "FAIL: want [5] got [", pdiff(x, y), "]"
    stop 1
end if
end subroutine show_diff
end program t
