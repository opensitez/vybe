! vybe-test: fortran/inquire_open_close_extended/ioc_unformatted_scratch_roundtrip
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: a, b
open(74, status='scratch', form='unformatted')
write(74) 8, 9
rewind(74)
read(74) a, b
close(74)
print *, a + b
end program t
