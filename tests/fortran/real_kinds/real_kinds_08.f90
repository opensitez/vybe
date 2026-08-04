! vybe-test: fortran/real_kinds/real_kinds_08
! origin: languages/fortran/tests/fortran/test_real_kinds.rs
program p
real(kind=8) :: a=1.5_8
print *, floor(a)
end program p
