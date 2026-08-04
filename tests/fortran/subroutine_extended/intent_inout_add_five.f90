! vybe-test: fortran/subroutine_extended/intent_inout_add_five
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: x
x = 10
call add_five(x)
if ((x) /= 15) then
    print *, "FAIL: want [15] got [", x, "]"
    stop 1
end if
contains
subroutine add_five(n)
integer, intent(inout) :: n
n = n + 5
end subroutine add_five
end program t
