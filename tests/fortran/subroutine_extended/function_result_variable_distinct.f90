! vybe-test: fortran/subroutine_extended/function_result_variable_distinct
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((incr(9)) /= 10) then
    print *, "FAIL: want [10] got [", incr(9), "]"
    stop 1
end if
contains
function incr(x) result(y)
integer, intent(in) :: x
integer :: y
y = x + 1
end function incr
end program t
