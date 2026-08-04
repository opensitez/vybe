! vybe-test: fortran/minmaxloc_val_extended/minloc_int_dip_at_index_two
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(6) = [8, 1, 6, 4, 9, 3]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 2) then
    print *, "FAIL: want [2] got [", loc(1), "]"
    stop 1
end if
if ((a(loc(1))) /= 1) then
    print *, "FAIL: want [1] got [", a(loc(1)), "]"
    stop 1
end if
end program t
