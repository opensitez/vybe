! vybe-test: fortran/namelist/nml_real_04
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
real::x=1.5
namelist /grp/ x
write(*,nml=grp)
end program p
