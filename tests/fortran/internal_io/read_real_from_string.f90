! vybe-test: fortran/internal_io/read_real_from_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=15) :: buf = '  3.14'
    real :: x
    read(buf, *) x
    print *, x
end program test
