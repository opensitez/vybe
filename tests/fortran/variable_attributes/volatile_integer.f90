! vybe-test: fortran/variable_attributes/volatile_integer
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    integer, volatile :: x = 0
    x = 42
    print *, x
end program test
