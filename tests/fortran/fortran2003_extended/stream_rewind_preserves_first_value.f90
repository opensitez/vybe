! vybe-test: fortran/fortran2003_extended/stream_rewind_preserves_first_value
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
integer :: first, second
open(12, status='scratch', access='stream', form='unformatted')
write(12) 100
write(12) 200
rewind(12)
read(12) first
read(12) second
close(12)
print *, first
print *, second
end program t
