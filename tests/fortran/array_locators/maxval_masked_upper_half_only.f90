! vybe-test: fortran/array_locators/maxval_masked_upper_half_only
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]
logical :: mask(8) = [.false., .false., .false., .false., .true., .true., .true., .true.]
if ((maxval(a, mask=mask)) /= 16) then
    print *, "FAIL: want [16] got [", maxval(a, mask=mask), "]"
    stop 1
end if
end program t
