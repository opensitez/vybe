! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_two_values_sum
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: a, b
open(21, status='scratch')
write(21, *) 10, 20
rewind(21)
read(21, *) a, b
close(21)
print *, a + b
end program t
