! vybe-test: fortran/stream_access_io/stream_unformatted_real_pair
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
real :: x, y
open(11, status='scratch', access='stream', form='unformatted')
write(11) 1.5, 2.5
rewind(11)
read(11) x, y
close(11)
print *, int(x + y)
end program t
