! vybe-test: fortran/internal_io/write_mixed_format
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=30) :: buf
    integer :: n = 10
    real :: x = 3.14
    write(buf, '(I4, F8.3)') n, x
    print *, trim(buf)
end program test
