! vybe-test: fortran/inquire_open_close_extended/ioc_newunit_write_read_integer
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: u, n
open(newunit=u, file='ioc_ext_new1.dat', status='replace')
write(u, '(I0)') 314
rewind(u)
read(u, '(I0)') n
close(u, status='delete')
print *, n
end program t
