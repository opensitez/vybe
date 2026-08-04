! vybe-test: fortran/array_semantics/arr_temp_07
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(3)=[1,2,3]
print *,a+1
end program p
