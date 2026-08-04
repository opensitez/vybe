! vybe-test: fortran/inquire_open_close_extended/ioc_newunit_list_directed
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: u, a, b
open(newunit=u, file='ioc_ext_new4.dat', status='replace')
write(u, *) 6, 7
rewind(u)
read(u, *) a, b
close(u, status='delete')
print *, a + b
end program t
