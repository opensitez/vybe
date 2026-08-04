! vybe-test: fortran/internal_io_extended/iio_read_formatted_i3_padded
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: buf = '007'
integer :: n
read(buf, '(I3)') n
if ((n) /= 7) then
    print *, "FAIL: want [7] got [", n, "]"
    stop 1
end if
end program t
