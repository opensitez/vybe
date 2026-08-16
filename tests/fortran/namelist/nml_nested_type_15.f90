! vybe-test: fortran/namelist/nml_nested_type_15
! origin: languages/fortran/tests/fortran/test_namelist.rs
program p
type::t
integer::x
end type t
type(t)::a(2)
namelist /grp/ a
write(*,nml=grp)
end program p
