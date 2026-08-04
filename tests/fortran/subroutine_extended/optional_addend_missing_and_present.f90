! vybe-test: fortran/subroutine_extended/optional_addend_missing_and_present
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((with_addend(5)) /= 5) then
    print *, "FAIL: want [5] got [", with_addend(5), "]"
    stop 1
end if
if ((with_addend(5, 3)) /= 8) then
    print *, "FAIL: want [8] got [", with_addend(5, 3), "]"
    stop 1
end if
contains
function with_addend(x, y) result(r)
integer, intent(in) :: x
integer, intent(in), optional :: y
integer :: r
if (present(y)) then
r = x + y
else
r = x
end if
end function with_addend
end program t
