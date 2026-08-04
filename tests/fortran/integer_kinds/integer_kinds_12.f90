! vybe-test: fortran/integer_kinds/integer_kinds_12
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer(kind=selected_int_kind(9)) :: x = 34
integer :: y
y = x + 1
print *, y
end program p
