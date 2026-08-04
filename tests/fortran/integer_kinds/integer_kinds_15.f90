! vybe-test: fortran/integer_kinds/integer_kinds_15
! origin: languages/fortran/tests/fortran/test_integer_kinds.rs
program p
integer(kind=8) :: a = 1_8
integer(kind=8) :: b = 2_8
print *, merge(a * b, -1, a == 1_8 .and. b == 2_8)
end program p
