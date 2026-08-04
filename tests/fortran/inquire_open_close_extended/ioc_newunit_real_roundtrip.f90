! vybe-test: fortran/inquire_open_close_extended/ioc_newunit_real_roundtrip
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: u
real :: r
open(newunit=u, file='ioc_ext_new5.dat', status='replace')
write(u, '(F0.1)') 2.5
rewind(u)
read(u, '(F0.1)') r
close(u, status='delete')
print *, int(r * 10)
end program t
