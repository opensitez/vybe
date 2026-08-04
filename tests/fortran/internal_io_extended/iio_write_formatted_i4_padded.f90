! vybe-test: fortran/internal_io_extended/iio_write_formatted_i4_padded
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: buf
write(buf, '(I4)') 7
print *, buf(1:4)
end program t
