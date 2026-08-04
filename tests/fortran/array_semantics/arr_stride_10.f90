! vybe-test: fortran/array_semantics/arr_stride_10
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(5)=[1,2,3,4,5]
print *,a(1:5:2)
end program p
