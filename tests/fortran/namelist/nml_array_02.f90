! vybe-test: fortran/namelist/nml_array_02
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::a(3)=[1,2,3]
namelist /grp/ a
write(*,nml=grp)
end program p
