! vybe-test: fortran/io_file_position/fio_stream_scratch_rewind_product
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: a, b
open(40, status='scratch', access='stream', form='unformatted')
write(40) 8, 9
rewind(40)
read(40) a, b
close(40)
print *, a * b
end program t
