! vybe-test: fortran/internal_io_extended/iio_read_logical_dot_true
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=12) :: buf = '.true.'
logical :: flag
read(buf, *) flag
if ((flag) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
end program t
