! vybe-test: fortran/coarrays/co_min_single_image
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: x = 7
    call co_min(x)
    print *, x
end program test
