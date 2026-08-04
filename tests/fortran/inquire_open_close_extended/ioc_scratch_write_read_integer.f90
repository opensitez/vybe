! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_write_read_integer
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(20, status='scratch')
write(20, '(I0)') 123
rewind(20)
read(20, '(I0)') n
close(20)
print *, n
end program t
