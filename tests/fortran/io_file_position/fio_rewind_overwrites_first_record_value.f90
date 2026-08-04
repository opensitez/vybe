! vybe-test: fortran/io_file_position/fio_rewind_overwrites_first_record_value
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: first
open(22, status='scratch')
write(22, '(I0)') 100
write(22, '(I0)') 200
rewind(22)
read(22, '(I0)') first
write(22, '(I0)') 300
rewind(22)
read(22, '(I0)') first
close(22)
print *, first
end program t
