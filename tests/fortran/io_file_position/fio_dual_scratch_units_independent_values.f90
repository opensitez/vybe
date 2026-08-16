! vybe-test: fortran/io_file_position/fio_dual_scratch_units_independent_values
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: a, b
open(15, status='scratch')
open(16, status='scratch')
write(15, '(I0)') 3
write(16, '(I0)') 7
rewind(15)
rewind(16)
read(15, *) a
read(16, *) b
close(15)
close(16)
print *, a
print *, b
end program t
