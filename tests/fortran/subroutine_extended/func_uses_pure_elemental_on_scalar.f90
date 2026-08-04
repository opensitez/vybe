! vybe-test: fortran/subroutine_extended/func_uses_pure_elemental_on_scalar
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((eincr(8)) /= 9) then
    print *, "FAIL: want [9] got [", eincr(8), "]"
    stop 1
end if
contains
elemental function eincr(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + 1
end function eincr
end program t
