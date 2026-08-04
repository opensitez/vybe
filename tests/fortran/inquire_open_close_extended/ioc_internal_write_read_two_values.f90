! vybe-test: fortran/inquire_open_close_extended/ioc_internal_write_read_two_values
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
character(len=30) :: buf
integer :: a, b
write(buf, *) 12, 34
read(buf, *) a, b
print *, a
print *, b
end program t
