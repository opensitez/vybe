! vybe-test: fortran/subroutines/function_returns_value
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program t
if ((double(5)) /= 10) then
    print *, "FAIL: want [10] got [", double(5), "]"
    stop 1
end if
contains
function double(x) result(res)
integer, intent(in) :: x
integer :: res
res = x * 2
end function double
end program t
