! vybe-test: fortran/internal_io_extended/iio_build_key_equals_value_line
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=16) :: buf
integer :: key = 7
write(buf, '(A,I0)') 'id=', key
print *, trim(buf)
end program t
