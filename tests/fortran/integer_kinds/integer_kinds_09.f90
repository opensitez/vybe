! vybe-test: fortran/integer_kinds/integer_kinds_09
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer(kind=8) :: a=1_8,b=2_8
print *, a+b
end program p
