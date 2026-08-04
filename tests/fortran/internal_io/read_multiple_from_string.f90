! vybe-test: fortran/internal_io/read_multiple_from_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf = '1 2 3'
    integer :: a, b, c
    read(buf, *) a, b, c
    print *, a + b + c
end program test
