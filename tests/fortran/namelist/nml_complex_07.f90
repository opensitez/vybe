! vybe-test: fortran/namelist/nml_complex_07
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
complex::z=(1.0,2.0)
namelist /grp/ z
write(*,nml=grp)
end program p
