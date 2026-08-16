! vybe-test: fortran/inquire_open_close_extended/ioc_iostat_zero_on_formatted_read
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n, ios
open(60, status='scratch')
write(60, '(I0)') 42
rewind(60)
read(60, *, iostat=ios) n
close(60)
print *, ios
print *, n
end program t
