! vybe-test: fortran/coarrays/co_reduce_sum
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x = this_image()
    call co_reduce(x, operator(+))
    if (this_image() == 1) print *, x
end program test
