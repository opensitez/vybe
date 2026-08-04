! vybe-test: fortran/namelist/nml_char_arr_16
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
character(len=3)::a(2)=(/'abc','def'/)
namelist /grp/ a
write(*,nml=grp)
end program p
