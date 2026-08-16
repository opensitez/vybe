! vybe-test: fortran/inquire_open_close_extended/ioc_internal_write_read_integer
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
character(len=20) :: buf
integer :: n
write(buf, '(I0)') 456
read(buf, *) n
print *, n
end program t
