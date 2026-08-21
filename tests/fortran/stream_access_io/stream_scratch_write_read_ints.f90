! vybe-test: fortran/stream_access_io/stream_scratch_write_read_ints
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
integer :: a, b
open(10, status='scratch', access='stream', form='unformatted')
write(10) 11, 22
rewind(10)
read(10) a, b
close(10)
print *, a + b
end program t
