! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_assigned_goto_multiple_targets
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_assigned_goto_multiple_targets
integer label
assign 10 to label
integer i
i = 1
go to label
if (trim('other') /= "matched") then
    print *, "FAIL: want [matched] got [", 'other', "]"
    stop 1
end if
10          print *, 'matched'
end program statement_f77_legacy_compat_assigned_goto_multiple_targets
