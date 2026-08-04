! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_explicit_type_suffix
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_explicit_type_suffix
integer*4 i
i = 1
if ((i) /= 1) then
    print *, "FAIL: want [1] got [", i, "]"
    stop 1
end if
end program statement_f77_legacy_compat_explicit_type_suffix
