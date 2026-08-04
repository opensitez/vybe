! vybe-test: fortran/fortran_catalog_calibrate/get_command
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
character(len=32) :: cmd
integer :: stat
stat = get_command(cmd)
print *, stat
end program t
