! vybe-test: fortran/internal_io_extended/iio_write_string_literal_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf
write(buf, *) 'data'
print *, trim(adjustl(buf))
end program t
