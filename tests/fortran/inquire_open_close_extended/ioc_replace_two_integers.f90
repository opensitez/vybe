! vybe-test: fortran/inquire_open_close_extended/ioc_replace_two_integers
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: a, b
open(42, file='ioc_ext_rep3.dat', status='replace')
write(42, *) 3, 4
rewind(42)
read(42, *) a, b
close(42, status='delete')
print *, a * b
end program t
