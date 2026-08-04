! vybe-test: fortran/minmaxloc_val_extended/minval_int_sparse_dips
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(7) = [100, 100, 3, 100, 100, 1, 100]
if ((minval(a)) /= 1) then
    print *, "FAIL: want [1] got [", minval(a), "]"
    stop 1
end if
end program t
