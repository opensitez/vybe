! vybe-test: fortran/internal_io_extended/iio_read_formatted_f52
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = ' 2.50'
real :: x
read(buf, '(F5.2)') x
if (abs((x) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", x, "]"
    stop 1
end if
end program t
