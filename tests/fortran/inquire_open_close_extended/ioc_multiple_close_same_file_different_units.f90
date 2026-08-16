! vybe-test: fortran/inquire_open_close_extended/ioc_multiple_close_same_file_different_units
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(77, file='ioc_ext_multi.dat', status='replace')
write(77, '(I0)') 91
close(77)
open(78, file='ioc_ext_multi.dat', status='old')
read(78, *) n
close(78, status='delete')
print *, n
end program t
