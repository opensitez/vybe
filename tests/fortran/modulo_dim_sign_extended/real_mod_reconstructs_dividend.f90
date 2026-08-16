! vybe-test: fortran/modulo_dim_sign_extended/real_mod_reconstructs_dividend
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: a=29.5, b=6.0, q, r
q = a / b
r = mod(a, b)
if ((merge(1, 0, q*b + r == a)) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, q*b + r == a), "]"
    stop 1
end if
end program t
