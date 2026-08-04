! vybe-test: fortran/internal_io_extended/iio_write_formatted_l1_false
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=4) :: buf
write(buf, '(L1)') .false.
print *, buf(1:1)
end program t
