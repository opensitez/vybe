! vybe-test: fortran/subroutine_extended/intent_out_single_assign
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: n
call fill(n)
if ((n) /= 99) then
    print *, "FAIL: want [99] got [", n, "]"
    stop 1
end if
contains
subroutine fill(x)
integer, intent(out) :: x
x = 99
end subroutine fill
end program t
