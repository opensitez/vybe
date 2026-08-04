! vybe-test: fortran/fortran_catalog_calibrate/all_true
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
logical :: m(4) = [.true., .true., .true., .true.]
print *, all(m)
end program t
