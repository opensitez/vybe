! vybe-test: fortran/namelist/nml_two_vars_03
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::a=1,b=2
namelist /grp/ a,b
write(*,nml=grp)
end program p
