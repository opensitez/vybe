! vybe-test: fortran/internal_io_extended/iio_read_formatted_skip_x_descriptor
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = 'ab12'
integer :: n
read(buf, '(2X,I2)') n
if ((n) /= 12) then
    print *, "FAIL: want [12] got [", n, "]"
    stop 1
end if
end program t
