! vybe-test: fortran/internal_io_extended/iio_write_advance_no_two_fields
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=8) :: buf = '        '
write(buf(1:4), '(I2)', advance='no') 4
write(buf(3:6), '(I2)', advance='no') 5
print *, buf(1:4)
end program t
