! vybe-test: fortran/inquire_open_close_extended/ioc_scratch_formatted_three_lines
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: x, y, z
open(23, status='scratch')
write(23, '(I0)') 1
write(23, '(I0)') 2
write(23, '(I0)') 3
rewind(23)
read(23, '(I0)') x
read(23, '(I0)') y
read(23, '(I0)') z
close(23)
print *, x
print *, y
print *, z
end program t
