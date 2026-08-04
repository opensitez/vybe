! vybe-test: fortran/internal_io_extended/iio_read_three_integers_sum
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=16) :: buf = '4 5 6'
integer :: a, b, c
read(buf, *) a, b, c
if ((a + b + c) /= 15) then
    print *, "FAIL: want [15] got [", a + b + c, "]"
    stop 1
end if
end program t
