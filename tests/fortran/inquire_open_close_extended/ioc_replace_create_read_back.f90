! vybe-test: fortran/inquire_open_close_extended/ioc_replace_create_read_back
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(40, file='ioc_ext_rep1.dat', status='replace')
write(40, '(I0)') 999
rewind(40)
read(40, '(I0)') n
close(40, status='delete')
print *, n
end program t
