! vybe-test: fortran/array_locators/findloc_mask_only_last_true_element
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [1, 1, 1, 1, 9]
logical :: mask(5) = [.false., .false., .false., .false., .true.]
integer :: loc(1)
loc = findloc(a, 9, mask=mask)
if ((loc(1)) /= 5) then
    print *, "FAIL: want [5] got [", loc(1), "]"
    stop 1
end if
end program t
