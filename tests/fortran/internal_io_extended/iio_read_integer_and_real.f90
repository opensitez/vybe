! vybe-test: fortran/internal_io_extended/iio_read_integer_and_real
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf = '3 2.5'
integer :: i
real :: x
read(buf, *) i, x
if ((i + int(x)) /= 5) then
    print *, "FAIL: want [5] got [", i + int(x), "]"
    stop 1
end if
end program t
