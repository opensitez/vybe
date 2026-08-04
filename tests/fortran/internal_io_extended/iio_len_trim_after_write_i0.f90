! vybe-test: fortran/internal_io_extended/iio_len_trim_after_write_i0
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf
write(buf, '(I0)') 404
print *, len_trim(buf)
end program t
