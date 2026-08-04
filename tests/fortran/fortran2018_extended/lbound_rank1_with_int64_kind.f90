! vybe-test: fortran/fortran2018_extended/lbound_rank1_with_int64_kind
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
use iso_fortran_env
integer :: a(5)
if ((lbound(a, 1, kind=int64)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(a, 1, kind=int64), "]"
    stop 1
end if
end program t
