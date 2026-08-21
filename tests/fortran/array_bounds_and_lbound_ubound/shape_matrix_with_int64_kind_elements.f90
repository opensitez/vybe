! vybe-test: fortran/array_bounds_and_lbound_ubound/shape_matrix_with_int64_kind_elements
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
use iso_fortran_env
integer :: m(2,5)
integer(int64) :: sh(2)
sh = shape(m, kind=int64)
if ((sh(1)) /= 2) then
    print *, "FAIL: want [2] got [", sh(1), "]"
    stop 1
end if
if ((sh(2)) /= 5) then
    print *, "FAIL: want [5] got [", sh(2), "]"
    stop 1
end if
end program t
