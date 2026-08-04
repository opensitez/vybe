! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_real_value_roundtrip
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
real :: r
open(24, status='scratch')
write(24, '(F0.1)') 3.5
rewind(24)
read(24, '(F0.1)') r
close(24)
print *, int(r * 10)
end program t
