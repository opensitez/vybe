! vybe-test: fortran/internal_io_extended/iio_write_real_literal_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=16) :: buf
write(buf, *) 2.5
print *, trim(adjustl(buf))
end program t
