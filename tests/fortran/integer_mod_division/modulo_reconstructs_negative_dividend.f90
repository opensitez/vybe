! vybe-test: fortran/integer_mod_division/modulo_reconstructs_negative_dividend
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: a = -29, b = 6, q, r
q = a / b
r = modulo(a, b)
if ((q * b + r) /= -23) then
    print *, "FAIL: want [-23] got [", q * b + r, "]"
    stop 1
end if
end program t
