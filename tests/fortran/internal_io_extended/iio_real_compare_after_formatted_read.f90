! vybe-test: fortran/internal_io_extended/iio_real_compare_after_formatted_read
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = ' 1.00'
real :: x
read(buf, '(F5.2)') x
if ((x == 1.0) .neqv. .true.) then
    print *, "FAIL: want [true] got [", x == 1.0, "]"
    stop 1
end if
end program t
