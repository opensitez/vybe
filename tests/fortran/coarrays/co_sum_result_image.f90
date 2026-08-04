! vybe-test: fortran/coarrays/co_sum_result_image
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    real :: x = 1.0
    call co_sum(x, result_image=1)
    if (this_image() == 1) print *, x
end program test
