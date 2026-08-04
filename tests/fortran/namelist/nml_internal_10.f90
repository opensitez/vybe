! vybe-test: fortran/namelist/nml_internal_10
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=1
character(len=50)::buf
namelist /grp/ x
write(buf,nml=grp)
print *, trim(buf)
end program p
