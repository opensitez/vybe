! vybe-test: fortran/namelist/nml_write_internal_19
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=3
character(len=50)::buf
namelist /grp/ x
write(buf,nml=grp)
end program p
