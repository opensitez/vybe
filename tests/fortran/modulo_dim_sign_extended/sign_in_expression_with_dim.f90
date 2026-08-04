! vybe-test: fortran/modulo_dim_sign_extended/sign_in_expression_with_dim
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((sign(dim(8,3), -1)) /= -5) then
    print *, "FAIL: want [-5] got [", sign(dim(8,3), -1), "]"
    stop 1
end if
end program t
