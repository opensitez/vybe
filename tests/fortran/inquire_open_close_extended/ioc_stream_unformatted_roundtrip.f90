! vybe-test: fortran/inquire_open_close_extended/ioc_stream_unformatted_roundtrip
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: v
open(75, file='ioc_ext_stream.dat', access='stream', form='unformatted', status='replace')
write(75) 64
rewind(75)
read(75) v
close(75, status='delete')
print *, v
end program t
