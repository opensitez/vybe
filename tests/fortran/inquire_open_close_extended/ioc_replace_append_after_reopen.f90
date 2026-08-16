! vybe-test: fortran/inquire_open_close_extended/ioc_replace_append_after_reopen
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(46, file='ioc_ext_rep7.dat', status='replace')
write(46, '(I0)') 10
close(46)
open(46, file='ioc_ext_rep7.dat', status='old', position='append')
write(46, '(I0)') 5
rewind(46)
read(46, *) n
close(46, status='delete')
print *, n
end program t
