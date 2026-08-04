! vybe-test: fortran/internal_io_extended/iio_write_formatted_o_octal
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=4) :: buf
write(buf, '(O4)') 10
print *, trim(buf)
end program t
