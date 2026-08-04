! vybe-test: fortran/internal_io/write_read_roundtrip_int
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf
    integer :: x = 12345, y
    write(buf, '(I10)') x
    read(buf, '(I10)') y
    print *, x == y
end program test
