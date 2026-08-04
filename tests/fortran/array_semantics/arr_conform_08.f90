! vybe-test: fortran/array_semantics/arr_conform_08
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(3)=[1,2,3],b(3)=[4,5,6]
print *,a+b
end program p
