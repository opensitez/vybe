! vybe-test: fortran/internal_io/read_formatted_int
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=10) :: buf = '   42'
    integer :: n
    read(buf, '(I5)') n
    print *, n
end program test
