! vybe-test: fortran/minmaxloc_val_extended/maxval_int_with_negatives
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [-10, -3, -7, -1, -5]
if ((maxval(a)) /= -1) then
    print *, "FAIL: want [-1] got [", maxval(a), "]"
    stop 1
end if
end program t
