! vybe-test: fortran/inquire_open_close_extended/ioc_replace_sequential_access
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(45, file='ioc_ext_rep6.dat', status='replace', access='sequential')
write(45, '(I0)') 66
rewind(45)
read(45, *) n
close(45, status='delete')
print *, n
end program t
