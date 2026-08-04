! vybe-test: fortran/array_semantics/arr_section_assign_16
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(4)=[1,2,3,4]
a(2:3)=0
print *,a
end program p
