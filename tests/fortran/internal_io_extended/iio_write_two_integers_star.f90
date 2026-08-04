! vybe-test: fortran/internal_io_extended/iio_write_two_integers_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=16) :: buf
write(buf, *) 8, 9
print *, index(buf, '8')
end program t
