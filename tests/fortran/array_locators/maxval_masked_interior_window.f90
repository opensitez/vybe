! vybe-test: fortran/array_locators/maxval_masked_interior_window
! origin: languages/fortran/tests/fortran/test_array_locators.rs
program t
integer :: a(7) = [100, 5, 20, 35, 50, 65, 200]
logical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]
if ((maxval(a, mask=mask)) /= 65) then
    print *, "FAIL: want [65] got [", maxval(a, mask=mask), "]"
    stop 1
end if
end program t
