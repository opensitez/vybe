! vybe-test: fortran/minmaxloc_val_extended/maxloc_mask_excludes_global_max
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [10, 20, 30, 40, 50]
logical :: mask(5) = [.true., .true., .true., .false., .false.]
integer :: loc(1)
loc = maxloc(a, mask=mask)
if ((loc(1)) /= 3) then
    print *, "FAIL: want [3] got [", loc(1), "]"
    stop 1
end if
if ((a(loc(1))) /= 30) then
    print *, "FAIL: want [30] got [", a(loc(1)), "]"
    stop 1
end if
end program t
