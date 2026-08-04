! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_assigned_goto
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_assigned_goto
integer label
assign 20 to label
goto label
if (trim('skip') /= "hit") then
    print *, "FAIL: want [hit] got [", 'skip', "]"
    stop 1
end if
20          print *, 'hit'
end program statement_f77_legacy_compat_assigned_goto
