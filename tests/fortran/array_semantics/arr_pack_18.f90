! vybe-test: fortran/array_semantics/arr_pack_18
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(3)=[1,2,3]
print *,pack(a,a>1)
end program p
