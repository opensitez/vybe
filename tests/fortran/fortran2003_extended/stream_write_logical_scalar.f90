! vybe-test: fortran/fortran2003_extended/stream_write_logical_scalar
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
logical :: flag
open(13, status='scratch', access='stream', form='unformatted')
write(13) .true.
rewind(13)
read(13) flag
close(13)
print *, flag
end program t
