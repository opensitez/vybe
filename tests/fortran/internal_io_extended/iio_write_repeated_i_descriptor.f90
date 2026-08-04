! vybe-test: fortran/internal_io_extended/iio_write_repeated_i_descriptor
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf
write(buf, '(2I4)') 11, 22
print *, buf(1:8)
end program t
