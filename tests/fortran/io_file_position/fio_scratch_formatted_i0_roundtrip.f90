! vybe-test: fortran/io_file_position/fio_scratch_formatted_i0_roundtrip
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: n
open(11, status='scratch')
write(11, '(I0)') 55
rewind(11)
read(11, *) n
close(11)
print *, n
end program t
