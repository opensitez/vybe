! vybe-test: fortran/array_semantics/arr_deferred_shape_04
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer, allocatable::a(:)
allocate(a(3))
end program p
