! vybe-test: fortran/internal_io_extended/iio_write_read_index_verification
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf
write(buf, '(A)') 'needle'
print *, index(buf, 'eed')
end program t
