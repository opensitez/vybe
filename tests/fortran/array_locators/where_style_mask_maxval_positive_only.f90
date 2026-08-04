! vybe-test: fortran/array_locators/where_style_mask_maxval_positive_only
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(5) = [3, -1, 5, -2, 4]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
if ((maxval(a, mask=mask)) /= 5) then
    print *, "FAIL: want [5] got [", maxval(a, mask=mask), "]"
    stop 1
end if
end program t
