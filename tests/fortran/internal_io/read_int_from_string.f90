! vybe-test: fortran/internal_io/read_int_from_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=10) :: buf = '   42'
    integer :: n
    read(buf, *) n
    print *, n
end program test
