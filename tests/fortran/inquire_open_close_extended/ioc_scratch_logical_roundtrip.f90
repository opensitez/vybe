! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_logical_roundtrip
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
logical :: flag
open(32, status='scratch')
write(32, *) .true.
rewind(32)
read(32, *) flag
close(32)
if (flag) then
print *, 1
else
print *, 0
end if
end program t
