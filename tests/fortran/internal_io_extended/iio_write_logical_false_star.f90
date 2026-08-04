! vybe-test: fortran/internal_io_extended/iio_write_logical_false_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=8) :: buf
write(buf, *) .false.
print *, trim(adjustl(buf))
end program t
