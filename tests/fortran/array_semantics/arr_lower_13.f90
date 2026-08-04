! vybe-test: fortran/array_semantics/arr_lower_13
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(-1:1)
print *,lbound(a)
end program p
