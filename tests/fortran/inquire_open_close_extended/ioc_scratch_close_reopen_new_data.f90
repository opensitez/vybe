! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_close_reopen_new_data
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: v
open(22, status='scratch')
write(22, '(I0)') 7
close(22)
open(22, status='scratch')
write(22, '(I0)') 42
rewind(22)
read(22, *) v
close(22)
print *, v
end program t
