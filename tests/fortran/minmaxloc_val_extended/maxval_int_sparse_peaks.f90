! vybe-test: fortran/minmaxloc_val_extended/maxval_int_sparse_peaks
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(7) = [0, 0, 15, 0, 0, 20, 0]
if ((maxval(a)) /= 20) then
    print *, "FAIL: want [20] got [", maxval(a), "]"
    stop 1
end if
end program t
