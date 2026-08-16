! vybe-test: fortran/internal_io_extended/iio_integer_arithmetic_after_parse
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=6) :: buf = '6'
integer :: n
read(buf, *) n
if ((n * n) /= 36) then
    print *, "FAIL: want [36] got [", n * n, "]"
    stop 1
end if
end program t
