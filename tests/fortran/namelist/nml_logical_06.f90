! vybe-test: fortran/namelist/nml_logical_06
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
logical::l=.true.
namelist /grp/ l
write(*,nml=grp)
end program p
