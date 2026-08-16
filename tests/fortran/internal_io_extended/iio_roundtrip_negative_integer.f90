! vybe-test: fortran/internal_io_extended/iio_roundtrip_negative_integer
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf
integer :: x = -18, y
write(buf, '(I0)') x
read(buf, *) y
print *, y
end program t
