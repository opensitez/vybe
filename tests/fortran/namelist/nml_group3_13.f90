! vybe-test: fortran/namelist/nml_group3_13
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=1
namelist /numbers/ x
write(*,nml=numbers)
end program p
