! vybe-test: fortran/array_locators/maxval_masked_skips_leading_nine
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(6) = [1, 9, 2, 8, 3, 7]
logical :: mask(6) = [.false., .false., .true., .true., .true., .true.]
if ((maxval(a, mask=mask)) /= 9) then
    print *, "FAIL: want [9] got [", maxval(a, mask=mask), "]"
    stop 1
end if
end program t
