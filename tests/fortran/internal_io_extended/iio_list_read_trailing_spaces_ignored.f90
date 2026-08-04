! vybe-test: fortran/internal_io_extended/iio_list_read_trailing_spaces_ignored
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = '99    '
integer :: n
read(buf, *) n
if ((n + 1) /= 100) then
    print *, "FAIL: want [100] got [", n + 1, "]"
    stop 1
end if
end program t
