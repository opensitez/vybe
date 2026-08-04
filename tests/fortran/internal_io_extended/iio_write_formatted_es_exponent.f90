! vybe-test: fortran/internal_io_extended/iio_write_formatted_es_exponent
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=14) :: buf
write(buf, '(ES9.2)') 0.0045
print *, len_trim(buf)
end program t
