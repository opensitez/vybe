! vybe-test: fortran/subroutine_extended/intent_inout_swap_scalars
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a, b
a = 3
b = 7
call swap_int(a, b)
if ((a) /= 7) then
    print *, "FAIL: want [7] got [", a, "]"
    stop 1
end if
if ((b) /= 3) then
    print *, "FAIL: want [3] got [", b, "]"
    stop 1
end if
contains
subroutine swap_int(x, y)
integer, intent(inout) :: x, y
integer :: t
t = x
x = y
y = t
end subroutine swap_int
end program t
