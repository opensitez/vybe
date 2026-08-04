! vybe-test: fortran/array_locators/where_style_mask_findloc_value_five
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [3, -1, 5, -2, 4]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: loc(1)
loc = findloc(a, 5, mask=mask)
if ((loc(1)) /= 3) then
    print *, "FAIL: want [3] got [", loc(1), "]"
    stop 1
end if
end program t
