! vybe-test: fortran/subroutine_extended/internal_three_level_sum
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call level_a()
contains
subroutine level_a()
call level_b(1)
end subroutine level_a
subroutine level_b(x)
integer, intent(in) :: x
call level_c(x, 2)
end subroutine level_b
subroutine level_c(a, b)
integer, intent(in) :: a, b
if ((a + b) /= 3) then
    print *, "FAIL: want [3] got [", a + b, "]"
    stop 1
end if
end subroutine level_c
end program t
