! vybe-test: fortran/array_locators/findloc_mask_all_true_finds_first
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(4) = [2, 3, 4, 5]
logical :: mask(4) = [.true., .true., .true., .true.]
integer :: loc(1)
loc = findloc(a, 3, mask=mask)
if ((loc(1)) /= 2) then
    print *, "FAIL: want [2] got [", loc(1), "]"
    stop 1
end if
end program t
