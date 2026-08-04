! vybe-test: fortran/inquire_open_close_extended/ioc_iostat_zero_on_list_write
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: ios
open(61, status='scratch')
write(61, *, iostat=ios) 1, 2, 3
close(61)
print *, ios
end program t
