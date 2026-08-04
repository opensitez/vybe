! vybe-test: fortran/subroutine_extended/intent_inout_bump_array_sum
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: a(3)
a = [1, 2, 3]
call bump_all(a)
if ((sum(a)) /= 9) then
    print *, "FAIL: want [9] got [", sum(a), "]"
    stop 1
end if
contains
subroutine bump_all(v)
integer, intent(inout) :: v(3)
v = v + 1
end subroutine bump_all
end program t
