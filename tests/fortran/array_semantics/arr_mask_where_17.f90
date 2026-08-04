! vybe-test: fortran/array_semantics/arr_mask_where_17
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(3)=[1,2,3]
where(a>1) a=a+1
print *,a
end program p
