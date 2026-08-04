! vybe-test: fortran/modulo_dim_sign_extended/sign_real_scaled_neg314
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(sign(3.14, -1.0)*100)) /= -314) then
    print *, "FAIL: want [-314] got [", nint(sign(3.14, -1.0)*100), "]"
    stop 1
end if
end program t
