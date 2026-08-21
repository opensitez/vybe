! vybe-test: fortran/enumerations/enum_auto_increment
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    enum, bind(c)
        enumerator :: NORTH, SOUTH, EAST, WEST
    end enum
    integer :: dir = EAST
    print *, dir
end program test
