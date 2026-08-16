! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_statement_function
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_statement_function
integer n
integer x
integer square
square(x) = x * x
n = 5
if ((square(n)) /= 25) then
    print *, "FAIL: want [25] got [", square(n), "]"
    stop 1
end if
if ((square(3)) /= 9) then
    print *, "FAIL: want [9] got [", square(3), "]"
    stop 1
end if
end program statement_f77_legacy_compat_statement_function
