! vybe-test: fortran/internal_io_extended/iio_read_formatted_two_i4_fields
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf = '  3  14'
integer :: a, b
read(buf, '(2I4)') a, b
if ((a + b) /= 17) then
    print *, "FAIL: want [17] got [", a + b, "]"
    stop 1
end if
end program t
