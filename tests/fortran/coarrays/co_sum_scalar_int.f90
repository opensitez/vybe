! vybe-test: fortran/coarrays/co_sum_scalar_int
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x
    x = this_image()
    call co_sum(x)
    if (this_image() == 1) print *, x
end program test
