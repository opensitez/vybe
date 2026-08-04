! vybe-test: fortran/inquire_open_close_extended/ioc_replace_formatted_record
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(43, file='ioc_ext_rep4.dat', status='replace', form='formatted')
write(43, '(I0)') 77
rewind(43)
read(43, '(I0)') n
close(43, status='delete')
print *, n
end program t
