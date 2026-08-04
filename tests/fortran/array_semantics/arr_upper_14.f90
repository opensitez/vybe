! vybe-test: fortran/array_semantics/arr_upper_14
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(2:4)
print *,ubound(a)
end program p
