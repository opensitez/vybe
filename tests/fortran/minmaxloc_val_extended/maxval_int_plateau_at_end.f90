! vybe-test: fortran/minmaxloc_val_extended/maxval_int_plateau_at_end
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(6) = [1, 2, 3, 8, 8, 8]
if ((maxval(a)) /= 8) then
    print *, "FAIL: want [8] got [", maxval(a), "]"
    stop 1
end if
end program t
