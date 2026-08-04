! vybe-test: fortran/coarrays/co_min_scalar
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x = 100 - this_image()
    call co_min(x)
    print *, x
end program test
