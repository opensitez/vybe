! vybe-test: fortran/fortran2003/enum_explicit_values
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    enum, bind(c)
        enumerator :: LOW = 1, MEDIUM = 5, HIGH = 10
    end enum
    integer :: level = HIGH
    print *, level
end program test
