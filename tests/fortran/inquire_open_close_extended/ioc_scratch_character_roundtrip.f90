! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_character_roundtrip
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
character(len=3) :: s
open(33, status='scratch')
write(33, '(A)') 'abc'
rewind(33)
read(33, '(A)') s
close(33)
print *, s
end program t
