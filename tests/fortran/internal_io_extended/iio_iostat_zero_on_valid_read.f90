! vybe-test: fortran/internal_io_extended/iio_iostat_zero_on_valid_read
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: buf = '88'
integer :: n, ios
read(buf, *, iostat=ios) n
if (ios == 0) print *, n
end program t
