! vybe-test: fortran/internal_io_extended/iio_write_formatted_b_binary
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=8) :: buf
write(buf, '(B8)') 15
print *, trim(buf)
end program t
