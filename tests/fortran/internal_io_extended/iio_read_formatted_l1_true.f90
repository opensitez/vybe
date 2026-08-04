! vybe-test: fortran/internal_io_extended/iio_read_formatted_l1_true
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=4) :: buf = 'T   '
logical :: flag
read(buf, '(L1)') flag
if ((flag) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
end program t
