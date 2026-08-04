! vybe-test: fortran/inquire_open_close_extended/ioc_close_status_keep_file
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(70, file='ioc_ext_keep.dat', status='replace')
write(70, '(I0)') 50
close(70, status='keep')
open(70, file='ioc_ext_keep.dat', status='old')
read(70, '(I0)') n
close(70, status='delete')
print *, n
end program t
