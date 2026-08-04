! vybe-test: fortran/fortran_catalog_calibrate/adjustl
! origin: languages/fortran/tests/fortran/fortran_catalog_calibrate.rs
program t
character(len=10) :: s = '   data'
print *, len_trim(adjustl(s))
end program t
