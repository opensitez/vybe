! vybe-test: fortran/fortran2018_extended/typeof_real_scalar
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    real :: x = 2.5
    print *, typeof(x)
end program t
