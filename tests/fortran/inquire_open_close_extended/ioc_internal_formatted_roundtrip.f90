! vybe-test: fortran/inquire_open_close_extended/ioc_internal_formatted_roundtrip
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
character(len=10) :: buf
real :: r
write(buf, '(F0.1)') 4.5
read(buf, '(F0.1)') r
print *, int(r * 10)
end program t
