! vybe-test: fortran/inquire_open_close_extended/ioc_position_asis_reopen
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(73, file='ioc_ext_asis.dat', status='replace')
write(73, '(I0)') 33
close(73)
open(73, file='ioc_ext_asis.dat', status='old', position='asis')
read(73, '(I0)') n
close(73, status='delete')
print *, n
end program t
