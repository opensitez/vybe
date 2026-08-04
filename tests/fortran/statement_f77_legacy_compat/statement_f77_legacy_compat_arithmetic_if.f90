! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_arithmetic_if
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_arithmetic_if
real x
x = -1.0
if (x) 10, 20, 30
10          print *, 'neg'
stop 0
20          print *, 'zero'
stop 0
30          print *, 'pos'
end program statement_f77_legacy_compat_arithmetic_if
