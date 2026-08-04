! vybe-test: fortran/namelist/nml_pointer_09
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
integer,target::x=1
integer,pointer::p
p=>x
namelist /grp/ p
write(*,nml=grp)
end program p
