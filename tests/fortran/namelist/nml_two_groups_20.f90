! vybe-test: fortran/namelist/nml_two_groups_20
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=1,y=2
namelist /g1/ x
namelist /g2/ y
write(*,nml=g1)
write(*,nml=g2)
end program p
