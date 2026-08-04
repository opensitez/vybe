! vybe-test: fortran/inquire_open_close_extended/ioc_iostat_write_after_close_fails
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: ios
open(66, status='scratch')
close(66)
write(66, '(I0)', iostat=ios) 1
if (ios /= 0) then
print *, 1
else
print *, 0
end if
end program t
