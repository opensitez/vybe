! vybe-test: fortran/internal_io_extended/iio_write_named_unit_and_fmt
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: buf
write(unit=buf, fmt='(I0)') 13
print *, trim(buf)
end program t
