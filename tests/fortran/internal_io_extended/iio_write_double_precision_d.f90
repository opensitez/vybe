! vybe-test: fortran/internal_io_extended/iio_write_double_precision_d
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=20) :: buf
real(kind=8) :: x = 6.25d0
write(buf, '(F6.2)') x
print *, trim(buf)
end program t
