! vybe-test: fortran/namelist/nml_group2_12
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=1
namelist /a/ x
write(*,nml=a)
end program p
