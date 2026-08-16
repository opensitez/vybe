! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_multiple_rewind_same_value
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(34, status='scratch')
write(34, '(I0)') 55
rewind(34)
read(34, *) n
rewind(34)
read(34, *) n
close(34)
print *, n
end program t
