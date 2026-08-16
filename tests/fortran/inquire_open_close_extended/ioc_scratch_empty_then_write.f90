! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_empty_then_write
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(35, status='scratch')
write(35, '(I0)') 0
rewind(35)
read(35, *) n
close(35)
print *, n
end program t
