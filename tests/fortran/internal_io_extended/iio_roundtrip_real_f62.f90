! vybe-test: fortran/internal_io_extended/iio_roundtrip_real_f62
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf
real :: x = 0.75, y
write(buf, '(F6.2)') x
read(buf, '(F6.2)') y
print *, y
end program t
