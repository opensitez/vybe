! vybe-test: fortran/fortran_catalog_calibrate/cpu_time
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
real :: t
call cpu_time(t)
print *, nint(t*100)
end program t
