! vybe-test: fortran/internal_io_extended/iio_write_formatted_f62
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=8) :: buf
write(buf, '(F6.2)') 1.25
print *, trim(buf)
end program t
