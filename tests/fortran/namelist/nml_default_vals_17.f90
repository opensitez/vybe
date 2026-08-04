! vybe-test: fortran/namelist/nml_default_vals_17
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=0
namelist /grp/ x
write(*,nml=grp)
end program p
