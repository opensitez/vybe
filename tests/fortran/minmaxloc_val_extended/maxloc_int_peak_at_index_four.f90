! vybe-test: fortran/minmaxloc_val_extended/maxloc_int_peak_at_index_four
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(6) = [2, 5, 3, 9, 1, 7]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 4) then
    print *, "FAIL: want [4] got [", loc(1), "]"
    stop 1
end if
if ((a(loc(1))) /= 9) then
    print *, "FAIL: want [9] got [", a(loc(1)), "]"
    stop 1
end if
end program t
