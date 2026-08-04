! vybe-test: fortran/internal_io_extended/iio_roundtrip_star_both_directions
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf
integer :: x = 55, y
write(buf, *) x
read(buf, *) y
print *, y
end program t
