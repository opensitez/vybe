! vybe-test: fortran/internal_io_extended/iio_write_formatted_a_width
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf
write(buf, '(A8)') 'Fortran'
print *, buf(1:8)
end program t
