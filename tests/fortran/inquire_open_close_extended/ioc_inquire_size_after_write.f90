! vybe-test: fortran/inquire_open_close_extended/ioc_inquire_size_after_write
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: sz
open(57, file='ioc_ext_size.dat', status='replace')
write(57, '(I0)') 12345
inquire(unit=57, size=sz)
close(57, status='delete')
if (sz > 0) then
print *, 1
else
print *, 0
end if
end program t
