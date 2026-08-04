! vybe-test: fortran/subroutine_extended/elemental_subroutine_negate_inout
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(3)
a = [5, -2, 7]
call enegate(a)
if ((sum(a)) /= -10) then
    print *, "FAIL: want [-10] got [", sum(a), "]"
    stop 1
end if
contains
elemental subroutine enegate(x)
integer, intent(inout) :: x
x = -x
end subroutine enegate
end program t
