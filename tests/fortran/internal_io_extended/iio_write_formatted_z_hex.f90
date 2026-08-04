! vybe-test: fortran/internal_io_extended/iio_write_formatted_z_hex
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: buf
write(buf, '(Z4)') 255
print *, trim(buf)
end program t
