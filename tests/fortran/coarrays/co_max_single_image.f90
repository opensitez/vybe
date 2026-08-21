! vybe-test: fortran/coarrays/co_max_single_image
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real :: x = 3.14
    call co_max(x)
    print *, x
end program test
