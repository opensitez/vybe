! vybe-test: fortran/minmaxloc_val_extended/minval_int_plateau_at_start
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(6) = [2, 2, 2, 5, 6, 7]
if ((minval(a)) /= 2) then
    print *, "FAIL: want [2] got [", minval(a), "]"
    stop 1
end if
end program t
