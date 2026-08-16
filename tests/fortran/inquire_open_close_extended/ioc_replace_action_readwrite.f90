! vybe-test: fortran/inquire_open_close_extended/ioc_replace_action_readwrite
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(44, file='ioc_ext_rep5.dat', status='replace', action='readwrite')
write(44, '(I0)') 88
rewind(44)
read(44, *) n
close(44, status='delete')
print *, n
end program t
