! vybe-test: fortran/fortran2018_extended/typeof_integer_scalar
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: x = 7
    print *, typeof(x)
end program t
