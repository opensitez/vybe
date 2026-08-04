! vybe-test: fortran/subroutine_extended/nested_function_call_chain
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((twice(twice(3))) /= 12) then
    print *, "FAIL: want [12] got [", twice(twice(3)), "]"
    stop 1
end if
contains
function twice(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * 2
end function twice
end program t
