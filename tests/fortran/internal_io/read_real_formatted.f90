! vybe-test: fortran/internal_io/read_real_formatted
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=15) :: buf = ' 3.14159'
    real :: x
    read(buf, '(F8.5)') x
    print *, x
end program test
