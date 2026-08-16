! vybe-test: fortran/inquire_open_close_extended/ioc_dual_scratch_independent
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: a, b
open(30, status='scratch')
open(31, status='scratch')
write(30, '(I0)') 11
write(31, '(I0)') 22
rewind(30)
rewind(31)
read(30, *) a
read(31, *) b
close(30)
close(31)
print *, a
print *, b
end program t
