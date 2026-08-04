! vybe-test: fortran/minmaxloc_val_extended/minloc_real_dip_index
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
real :: a(4) = [1.0, 4.0, 2.5, 3.0]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
