! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_hollerith_literal
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_hollerith_literal
character*4 c
c = 4hABCD
if (trim(c) /= "ABCD") then
    print *, "FAIL: want [ABCD] got [", c, "]"
    stop 1
end if
end program statement_f77_legacy_compat_hollerith_literal
