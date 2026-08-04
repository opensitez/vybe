! vybe-test: fortran/fortran2018_extended/size_matrix_dim2_with_int64_kind
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
use iso_fortran_env
integer :: m(3,4)
if ((size(m, 2, kind=int64)) /= 4) then
    print *, "FAIL: want [4] got [", size(m, 2, kind=int64), "]"
    stop 1
end if
end program t
