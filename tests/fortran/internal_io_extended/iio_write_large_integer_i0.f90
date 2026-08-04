! vybe-test: fortran/internal_io_extended/iio_write_large_integer_i0
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf
write(buf, '(I0)') 987654321
print *, trim(buf)
end program t
