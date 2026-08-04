! vybe-test: fortran/minmaxloc_val_extended/minloc_mask_excludes_global_min
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [10, 20, 30, 40, 50]
logical :: mask(5) = [.false., .true., .true., .true., .true.]
integer :: loc(1)
loc = minloc(a, mask=mask)
if ((loc(1)) /= 2) then
    print *, "FAIL: want [2] got [", loc(1), "]"
    stop 1
end if
if ((a(loc(1))) /= 20) then
    print *, "FAIL: want [20] got [", a(loc(1)), "]"
    stop 1
end if
end program t
