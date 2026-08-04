! vybe-test: fortran/array_semantics/arr_reshape_15
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(2,2)
a=reshape([1,2,3,4],[2,2])
print *,a
end program p
