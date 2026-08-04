! vybe-test: fortran/coarrays/co_max_real
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    real :: r = 3.14 * this_image()
    call co_max(r, result_image=1)
    if (this_image() == 1) print *, r
end program test
