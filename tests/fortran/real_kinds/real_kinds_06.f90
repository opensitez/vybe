! vybe-test: fortran/real_kinds/real_kinds_06
! origin: languages/fortran/tests/fortran/test_real_kinds.rs
program p
real(kind=4) :: a=1.0_4,b=2.0_4
print *, a+b
end program p
