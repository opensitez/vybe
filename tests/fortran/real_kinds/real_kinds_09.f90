! vybe-test: fortran/real_kinds/real_kinds_09
! origin: languages/fortran/tests/fortran/test_real_kinds.rs
program p
real(kind=8) :: a=1.5_8
print *, ceiling(a)
end program p
