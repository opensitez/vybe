! vybe-test: fortran/internal_io_extended/iio_read_zero_value_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: buf = '0'
integer :: n
read(buf, *) n
if ((n) /= 0) then
    print *, "FAIL: want [0] got [", n, "]"
    stop 1
end if
end program t
