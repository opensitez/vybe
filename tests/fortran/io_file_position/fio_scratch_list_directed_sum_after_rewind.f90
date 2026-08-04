! vybe-test: fortran/io_file_position/fio_scratch_list_directed_sum_after_rewind
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: a, b
open(10, status='scratch')
write(10, *) 10, 20
rewind(10)
read(10, *) a, b
close(10)
print *, a + b
end program t
