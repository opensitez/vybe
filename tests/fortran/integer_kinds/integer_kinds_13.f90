! vybe-test: fortran/integer_kinds/integer_kinds_13
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer(kind=selected_int_kind(10)) :: x = 0
print *, x
end program p
