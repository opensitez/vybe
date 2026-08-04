! vybe-test: fortran/io_advanced/inquire_unit
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    logical :: opened
    inquire(unit=10, opened=opened)
    print *, opened
end program test
