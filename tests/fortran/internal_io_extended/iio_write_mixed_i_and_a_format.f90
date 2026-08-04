! vybe-test: fortran/internal_io_extended/iio_write_mixed_i_and_a_format
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf
write(buf, '(I0,A,I0)') 3, '-', 7
print *, trim(buf)
end program t
