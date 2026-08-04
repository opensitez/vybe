! vybe-test: fortran/internal_io_extended/iio_read_logical_dot_false
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf = '.false.'
logical :: flag
read(buf, *) flag
if ((flag) .neqv. .false.) then
    print *, "FAIL: want [false] got [", flag, "]"
    stop 1
end if
end program t
