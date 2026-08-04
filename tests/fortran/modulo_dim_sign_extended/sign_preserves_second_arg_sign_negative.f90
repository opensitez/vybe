! vybe-test: fortran/modulo_dim_sign_extended/sign_preserves_second_arg_sign_negative
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((sign(100, -3)) /= -100) then
    print *, "FAIL: want [-100] got [", sign(100, -3), "]"
    stop 1
end if
end program t
