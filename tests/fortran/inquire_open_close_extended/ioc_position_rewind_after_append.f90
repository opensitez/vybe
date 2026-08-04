! vybe-test: fortran/inquire_open_close_extended/ioc_position_rewind_after_append
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: n
open(72, file='ioc_ext_pos.dat', status='replace')
write(72, '(I0)') 100
close(72)
open(72, file='ioc_ext_pos.dat', status='old', position='append')
write(72, '(I0)') 200
rewind(72)
read(72, '(I0)') n
close(72, status='delete')
print *, n
end program t
