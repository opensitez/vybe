! vybe-test: fortran/internal_io_extended/iio_write_negative_int_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf
write(buf, *) -99
print *, trim(adjustl(buf))
end program t
