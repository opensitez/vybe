! vybe-test: fortran/fortran2003/enum_basic
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    enum, bind(c)
        enumerator :: RED = 0, GREEN = 1, BLUE = 2
    end enum
    integer :: color = GREEN
    print *, color
end program test
