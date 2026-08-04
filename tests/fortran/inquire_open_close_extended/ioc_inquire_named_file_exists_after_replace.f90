! vybe-test: fortran/inquire_open_close_extended/ioc_inquire_named_file_exists_after_replace
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
logical :: exists
open(53, file='ioc_ext_exist.dat', status='replace')
write(53, '(I0)') 1
close(53)
inquire(file='ioc_ext_exist.dat', exist=exists)
open(53, file='ioc_ext_exist.dat', status='old')
close(53, status='delete')
if (exists) then
print *, 1
else
print *, 0
end if
end program t
