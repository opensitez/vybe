! vybe-test: fortran/fortran_catalog_calibrate/date_and_time
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
integer :: dt(8)
call date_and_time(values=dt)
print *, dt(1)
end program t
