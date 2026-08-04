! vybe-test: fortran/real_kinds/real_kinds_07
! origin: languages/fortran/tests/fortran/test_real_kinds.rs
program p
real(kind=8) :: a=4.0_8
print *, sqrt(a)
end program p
