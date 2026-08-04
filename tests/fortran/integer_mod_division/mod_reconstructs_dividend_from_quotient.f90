! vybe-test: fortran/integer_mod_division/mod_reconstructs_dividend_from_quotient
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: a = 29, b = 6, q, r
q = a / b
r = mod(a, b)
if ((q * b + r) /= 29) then
    print *, "FAIL: want [29] got [", q * b + r, "]"
    stop 1
end if
end program t
