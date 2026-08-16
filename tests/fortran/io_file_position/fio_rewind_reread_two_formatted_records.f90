! vybe-test: fortran/io_file_position/fio_rewind_reread_two_formatted_records
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: a, b
open(20, status='scratch')
write(20, '(I0)') 1
write(20, '(I0)') 2
rewind(20)
read(20, *) a
read(20, *) b
close(20)
print *, a
print *, b
end program t
