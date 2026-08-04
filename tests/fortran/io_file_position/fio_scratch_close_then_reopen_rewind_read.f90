! vybe-test: fortran/io_file_position/fio_scratch_close_then_reopen_rewind_read
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: v
open(12, status='scratch')
write(12, '(I0)') 9
close(12)
open(12, status='scratch')
write(12, '(I0)') 4
rewind(12)
read(12, '(I0)') v
close(12)
print *, v
end program t
