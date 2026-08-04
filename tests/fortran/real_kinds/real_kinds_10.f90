! vybe-test: fortran/real_kinds/real_kinds_10
! origin: languages/fortran/tests/fortran/test_real_kinds.rs
program p
real(kind=8) :: a=1.0_8,b=2.0_8
print *, nearest(a,b)
end program p
