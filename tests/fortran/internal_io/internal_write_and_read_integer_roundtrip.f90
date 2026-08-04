! vybe-test: fortran/internal_io/internal_write_and_read_integer_roundtrip
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=10) :: buf
    integer :: n
    write(buf, '(I0)') 87
    read(buf, '(I0)') n
    print *, n
end program test
