! vybe-test: fortran/inquire_open_close_extended/ioc_replace_delete_on_close
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(47, file='ioc_ext_rep8.dat', status='replace')
write(47, '(I0)') 44
close(47, status='delete')
print *, 44
end program t
