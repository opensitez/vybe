! vybe-test: fortran/io_file_position/fio_replace_discards_prior_session_content
! origin: languages/fortran/tests/fortran/test_io_file_position.rs
program t
integer :: n
open(31, file='fio_replace_stale.dat', status='replace')
write(31, '(I0)') 111
close(31)
open(31, file='fio_replace_stale.dat', status='replace')
write(31, '(I0)') 222
rewind(31)
read(31, *) n
close(31, status='delete')
print *, n
end program t
