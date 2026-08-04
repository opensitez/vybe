! vybe-test: fortran/integer_kinds/integer_kinds_10
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer(kind=4) :: a=7_4
print *, mod(a,3_4)
end program p
