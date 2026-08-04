! vybe-test: fortran/array_semantics/arr_zero_size_11
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer, allocatable::a(:)
allocate(a(0))
print *,size(a)
end program p
