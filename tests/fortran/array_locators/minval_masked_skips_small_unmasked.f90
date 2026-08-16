! vybe-test: fortran/array_locators/minval_masked_skips_small_unmasked
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(6) = [10, 1, 20, 2, 30, 3]
logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
if ((minval(a, mask=mask)) /= 10) then
    print *, "FAIL: want [10] got [", minval(a, mask=mask), "]"
    stop 1
end if
end program t
