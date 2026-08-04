! vybe-test: fortran/array_locators/maxval_masked_even_positions
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(6) = [11, 22, 33, 44, 55, 66]
logical :: mask(6) = [.false., .true., .false., .true., .false., .true.]
if ((maxval(a, mask=mask)) /= 66) then
    print *, "FAIL: want [66] got [", maxval(a, mask=mask), "]"
    stop 1
end if
end program t
