! vybe-test: fortran/minmaxloc_val_extended/minval_int_ascending_sequence
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
if ((minval(a)) /= 1) then
    print *, "FAIL: want [1] got [", minval(a), "]"
    stop 1
end if
end program t
