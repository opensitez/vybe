! vybe-test: fortran/inquire_open_close_extended/ioc_internal_character_preservation
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
character(len=10) :: buf, s
s = 'hello'
write(buf, '(A)') s
read(buf, '(A)') s
print *, s
end program t
