! vybe-test: fortran/internal_io_extended/iio_read_real_literal_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = '3.5'
real :: x
read(buf, *) x
if (abs((x) - 3.5) > 1.0e-6) then
    print *, "FAIL: want [3.5] got [", x, "]"
    stop 1
end if
end program t
