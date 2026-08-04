! vybe-test: fortran/integer_kinds/integer_kinds_08
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer(kind=4), parameter :: x=1_4
print *, x
end program p
