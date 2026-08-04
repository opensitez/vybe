! vybe-test: fortran/internal_io_extended/iio_read_scientific_notation_star
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf = '1.5e2'
real :: x
read(buf, *) x
if ((int(x)) /= 150) then
    print *, "FAIL: want [150] got [", int(x), "]"
    stop 1
end if
end program t
