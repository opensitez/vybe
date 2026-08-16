! vybe-test: fortran/inquire_open_close_extended/ioc_newunit_two_handles
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: u1, u2, a, b
open(newunit=u1, file='ioc_ext_new2a.dat', status='replace')
open(newunit=u2, file='ioc_ext_new2b.dat', status='replace')
write(u1, '(I0)') 10
write(u2, '(I0)') 20
rewind(u1)
rewind(u2)
read(u1, *) a
read(u2, *) b
close(u1, status='delete')
close(u2, status='delete')
print *, a + b
end program t
