! vybe-test: fortran/internal_io_extended/iio_write_zero_i0_descriptor
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=4) :: buf
write(buf, '(I0)') 0
print *, trim(buf)
end program t
