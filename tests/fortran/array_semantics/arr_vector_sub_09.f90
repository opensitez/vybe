! vybe-test: fortran/array_semantics/arr_vector_sub_09
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(4)=[1,2,3,4],i(2)=[1,3]
print *,a(i)
end program p
