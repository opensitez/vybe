! vybe-test: fortran/array_semantics/arr_assign_06
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(3),b(3)
a=[1,2,3]
b=a
print *,b
end program p
