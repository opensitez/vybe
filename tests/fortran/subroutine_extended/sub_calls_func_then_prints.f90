! vybe-test: fortran/subroutine_extended/sub_calls_func_then_prints
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call emit_square(5)
contains
function sq(n) result(r)
integer, intent(in) :: n
integer :: r
r = n * n
end function sq
subroutine emit_square(n)
integer, intent(in) :: n
if ((sq(n)) /= 25) then
    print *, "FAIL: want [25] got [", sq(n), "]"
    stop 1
end if
end subroutine emit_square
end program t
