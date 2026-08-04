! vybe-test: fortran/inquire_open_close_extended/ioc_iostat_endfile_detection
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n, ios
open(62, status='scratch')
write(62, '(I0)') 1
rewind(62)
read(62, '(I0)', iostat=ios) n
read(62, '(I0)', iostat=ios) n
close(62)
print *, ios
end program t
