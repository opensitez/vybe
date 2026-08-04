! vybe-test: fortran/internal_io_extended/iio_read_negative_integer_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=8) :: buf = '-42'
integer :: n
read(buf, *) n
if ((n) /= -42) then
    print *, "FAIL: want [-42] got [", n, "]"
    stop 1
end if
end program t
