! vybe-test: fortran/modulo_dim_sign_extended/dim_int_equal_operands
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((dim(5, 5)) /= 0) then
    print *, "FAIL: want [0] got [", dim(5, 5), "]"
    stop 1
end if
end program t
