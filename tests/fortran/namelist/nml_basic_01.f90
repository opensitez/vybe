! vybe-test: fortran/namelist/nml_basic_01
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=1
namelist /grp/ x
write(*,nml=grp)
end program p
