! vybe-test: fortran/namelist/nml_char_05
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
character(len=5)::s='abc'
namelist /grp/ s
write(*,nml=grp)
end program p
