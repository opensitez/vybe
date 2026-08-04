! vybe-test: fortran/namelist/nml_multi_arr_14
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::a(2)=[1,2],b(2)=[3,4]
namelist /grp/ a,b
write(*,nml=grp)
end program p
