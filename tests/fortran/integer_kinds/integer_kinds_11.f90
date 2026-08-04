! vybe-test: fortran/integer_kinds/integer_kinds_11
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer(kind=selected_int_kind(4)) :: x = 12
print *, x
end program p
