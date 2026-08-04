! vybe-test: fortran/namelist/nml_read_internal_18
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer::x=0
character(len=50)::buf='&grp x=7 /'
namelist /grp/ x
read(buf,nml=grp)
end program p
