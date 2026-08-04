! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_logical_and_goto_label_references
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_logical_and_goto_label_references
integer i
i = 1
if (i .gt. 0) goto 100
if (trim('bad') /= "good") then
    print *, "FAIL: want [good] got [", 'bad', "]"
    stop 1
end if
100         print *, 'good'
end program statement_f77_legacy_compat_logical_and_goto_label_references
