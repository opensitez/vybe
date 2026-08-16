! vybe-test: fortran/io_file_position/fio_rewind_after_endfile_then_append
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: tail
open(21, status='scratch')
write(21, '(I0)') 5
endfile(21)
rewind(21)
write(21, '(I0)') 6
rewind(21)
read(21, *) tail
close(21)
print *, tail
end program t
