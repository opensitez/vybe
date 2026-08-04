! vybe-test: fortran/inquire_open_close_extended/ioc_inquire_formatted_on_replace_file
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
character(len=20) :: frm
open(55, file='ioc_ext_form.dat', status='replace', form='formatted')
write(55, '(I0)') 1
inquire(unit=55, form=frm)
close(55, status='delete')
print *, frm(1:9)
end program t
