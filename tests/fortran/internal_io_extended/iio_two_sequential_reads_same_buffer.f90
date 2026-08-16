! vybe-test: fortran/internal_io_extended/iio_two_sequential_reads_same_buffer
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
character(len=10) :: buf = '2 3'
integer :: a, b
read(buf, *) a
read(buf, *) b
if ((a + b) /= 4) then
    print *, "FAIL: want [4] got [", a + b, "]"
    stop 1
end if
end program t
