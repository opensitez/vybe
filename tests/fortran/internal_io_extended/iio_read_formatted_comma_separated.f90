! vybe-test: fortran/internal_io_extended/iio_read_formatted_comma_separated
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = '2,4'
integer :: a, b
read(buf, '(I0,",",I0)') a, b
if ((a * b) /= 8) then
    print *, "FAIL: want [8] got [", a * b, "]"
    stop 1
end if
end program t
