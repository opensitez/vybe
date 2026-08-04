! vybe-test: fortran/internal_io/write_read_roundtrip_real
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf
    real :: x = 3.14, y
    write(buf, '(F10.4)') x
    read(buf, '(F10.4)') y
    print *, abs(x - y) < 1e-3
end program test
