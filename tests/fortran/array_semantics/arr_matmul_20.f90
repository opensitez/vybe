! vybe-test: fortran/array_semantics/arr_matmul_20
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(2,2)=reshape([1,2,3,4],[2,2]),b(2,2)=reshape([1,0,0,1],[2,2])
print *,matmul(a,b)
end program p
