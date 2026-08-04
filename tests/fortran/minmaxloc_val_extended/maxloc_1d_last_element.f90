! vybe-test: fortran/minmaxloc_val_extended/maxloc_1d_last_element
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(4) = [1, 2, 3, 99]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 4) then
    print *, "FAIL: want [4] got [", loc(1), "]"
    stop 1
end if
end program t
