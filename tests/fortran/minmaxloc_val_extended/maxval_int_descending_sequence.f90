! vybe-test: fortran/minmaxloc_val_extended/maxval_int_descending_sequence
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [9, 7, 5, 3, 1]
if ((maxval(a)) /= 9) then
    print *, "FAIL: want [9] got [", maxval(a), "]"
    stop 1
end if
end program t
