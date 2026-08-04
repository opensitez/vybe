! vybe-test: fortran/array_semantics/arr_bounds_12
! origin: languages/fortran/tests/fortran/test_array_semantics.rs
program p
integer::a(0:2)
print *,lbound(a),ubound(a)
end program p
