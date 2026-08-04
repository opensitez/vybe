! vybe-test: fortran/integer_kinds/integer_kinds_14
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer :: k
k = selected_int_kind(15)
print *, k >= 1 .or. k == -1
end program p
