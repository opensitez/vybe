! vybe-test: fortran/io_file_position/fio_replace_single_integer_roundtrip
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: n
open(30, file='fio_replace_one.dat', status='replace')
write(30, '(I0)') 88
rewind(30)
read(30, *) n
close(30, status='delete')
print *, n
end program t
