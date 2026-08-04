! vybe-test: fortran/array_locators/findloc_mask_middle_true_finds_center
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [1, 2, 3, 2, 1]
logical :: mask(5) = [.false., .false., .true., .false., .false.]
integer :: loc(1)
loc = findloc(a, 3, mask=mask)
if ((loc(1)) /= 3) then
    print *, "FAIL: want [3] got [", loc(1), "]"
    stop 1
end if
end program t
