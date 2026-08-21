! vybe-test: fortran/out_of_range/out_of_range_real_to_integer
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real :: x = 3.14
    print *, out_of_range(x, 0)
end program test
