! vybe-test: fortran/array_locators/minval_masked_upper_half_only
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(8) = [2, 4, 6, 8, 10, 12, 14, 16]
logical :: mask(8) = [.false., .false., .false., .false., .true., .true., .true., .true.]
if ((minval(a, mask=mask)) /= 10) then
    print *, "FAIL: want [10] got [", minval(a, mask=mask), "]"
    stop 1
end if
end program t
