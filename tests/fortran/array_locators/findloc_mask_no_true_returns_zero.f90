! vybe-test: fortran/array_locators/findloc_mask_no_true_returns_zero
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(3) = [1, 2, 3]
logical :: mask(3) = [.false., .false., .false.]
integer :: loc(1)
loc = findloc(a, 2, mask=mask)
if ((loc(1)) /= 0) then
    print *, "FAIL: want [0] got [", loc(1), "]"
    stop 1
end if
end program t
