! vybe-test: fortran/minmaxloc_val_extended/minval_int_with_negatives
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [-10, -3, -7, -1, -5]
if ((minval(a)) /= -10) then
    print *, "FAIL: want [-10] got [", minval(a), "]"
    stop 1
end if
end program t
