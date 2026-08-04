! vybe-test: fortran/namelist/nml_read_11
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x
character(len=50)::buf='&grp x=1 /'
namelist /grp/ x
read(buf,nml=grp)
print *, x
end program p
