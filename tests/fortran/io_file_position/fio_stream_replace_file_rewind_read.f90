! vybe-test: fortran/io_file_position/fio_stream_replace_file_rewind_read
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: v
open(41, file='fio_stream_pos.dat', access='stream', form='unformatted', status='replace')
write(41) 64
rewind(41)
read(41) v
close(41, status='delete')
print *, v
end program t
