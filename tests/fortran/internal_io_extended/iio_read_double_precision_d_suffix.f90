! vybe-test: fortran/internal_io_extended/iio_read_double_precision_d_suffix
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf = '2.0d1'
real(kind=8) :: d
read(buf, *) d
if ((int(d)) /= 20) then
    print *, "FAIL: want [20] got [", int(d), "]"
    stop 1
end if
end program t
