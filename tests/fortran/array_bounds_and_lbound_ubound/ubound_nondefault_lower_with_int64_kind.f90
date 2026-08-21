! vybe-test: fortran/array_bounds_and_lbound_ubound/ubound_nondefault_lower_with_int64_kind
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
use iso_fortran_env
integer :: a(2:6)
if ((ubound(a, 1, kind=int64)) /= 6) then
    print *, "FAIL: want [6] got [", ubound(a, 1, kind=int64), "]"
    stop 1
end if
end program t
