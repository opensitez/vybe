! vybe-test: fortran/array_locators/findloc_mask_skips_leading_positions
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [9, 9, 9, 4, 5]
logical :: mask(5) = [.false., .false., .false., .true., .true.]
integer :: loc(1)
loc = findloc(a, 4, mask=mask)
if ((loc(1)) /= 4) then
    print *, "FAIL: want [4] got [", loc(1), "]"
    stop 1
end if
end program t
