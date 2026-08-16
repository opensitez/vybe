! vybe-test: fortran/inquire_open_close_extended/ioc_close_unit_without_status
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(76, file='ioc_ext_plain.dat', status='replace')
write(76, '(I0)') 17
close(76)
open(76, file='ioc_ext_plain.dat', status='old')
read(76, *) n
close(76, status='delete')
print *, n
end program t
