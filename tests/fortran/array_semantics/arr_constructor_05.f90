! vybe-test: fortran/array_semantics/arr_constructor_05
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(3)
a=[1,2,3]
print *,a
end program p
