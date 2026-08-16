! vybe-test: fortran/inquire_open_close_extended/ioc_replace_truncates_old_content
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(41, file='ioc_ext_rep2.dat', status='replace')
write(41, '(I0)') 111
close(41)
open(41, file='ioc_ext_rep2.dat', status='replace')
write(41, '(I0)') 222
rewind(41)
read(41, *) n
close(41, status='delete')
print *, n
end program t
