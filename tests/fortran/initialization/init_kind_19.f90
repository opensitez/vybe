! vybe-test: fortran/initialization/init_kind_19
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer(kind=8)::x=1_8
print *,x
end program p
