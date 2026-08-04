! vybe-test: fortran/subroutine_extended/internal_sub_invokes_local_function
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call report_square(6)
contains
function square(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * n
end function square
subroutine report_square(n)
integer, intent(in) :: n
if ((square(n)) /= 36) then
    print *, "FAIL: want [36] got [", square(n), "]"
    stop 1
end if
end subroutine report_square
end program t
