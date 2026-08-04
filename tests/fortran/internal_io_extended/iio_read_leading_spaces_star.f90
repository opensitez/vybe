! vybe-test: fortran/internal_io_extended/iio_read_leading_spaces_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = '    17'
integer :: n
read(buf, *) n
if ((n) /= 17) then
    print *, "FAIL: want [17] got [", n, "]"
    stop 1
end if
end program t
