! vybe-test: fortran/if_construct_extended/arith_if_integer_expression_zero
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
integer :: n = 0
if (n) 10, 20, 30
10 print *, "neg"; goto 99
20 print *, "zer"; goto 99
30 print *, "pos"
99 continue
end program t
