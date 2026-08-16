! vybe-test: fortran/namelist/nml_derived_08
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
type::t
integer::x
end type t
type(t)::v
namelist /grp/ v
write(*,nml=grp)
end program p
