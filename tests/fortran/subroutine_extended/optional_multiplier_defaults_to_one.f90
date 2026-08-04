! vybe-test: fortran/subroutine_extended/optional_multiplier_defaults_to_one
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((scale_val(6)) /= 6) then
    print *, "FAIL: want [6] got [", scale_val(6), "]"
    stop 1
end if
if ((scale_val(6, 4)) /= 24) then
    print *, "FAIL: want [24] got [", scale_val(6, 4), "]"
    stop 1
end if
contains
function scale_val(x, factor) result(r)
integer, intent(in) :: x
integer, intent(in), optional :: factor
integer :: r
if (present(factor)) then
r = x * factor
else
r = x
end if
end function scale_val
end program t
