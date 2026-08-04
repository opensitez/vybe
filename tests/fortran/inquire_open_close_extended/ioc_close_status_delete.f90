! vybe-test: fortran/inquire_open_close_extended/ioc_close_status_delete
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
open(71, file='ioc_ext_del.dat', status='replace')
write(71, '(I0)') 1
close(71, status='delete')
print *, 1
end program t
