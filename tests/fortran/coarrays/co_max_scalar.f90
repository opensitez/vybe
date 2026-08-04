! vybe-test: fortran/coarrays/co_max_scalar
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x
    x = this_image() * 10
    call co_max(x)
    print *, x
end program test
