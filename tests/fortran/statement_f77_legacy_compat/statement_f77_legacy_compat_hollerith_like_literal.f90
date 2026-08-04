! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_hollerith_like_literal
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_hollerith_like_literal
character*12 c
data c / 'HELLOWORLD  ' /
if (trim(trim(c)) /= "HELLOWORLD") then
    print *, "FAIL: want [HELLOWORLD] got [", trim(c), "]"
    stop 1
end if
end program statement_f77_legacy_compat_hollerith_like_literal
