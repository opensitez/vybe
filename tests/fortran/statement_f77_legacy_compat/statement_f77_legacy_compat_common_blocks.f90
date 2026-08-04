! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_common_blocks
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_common_blocks
integer a, b
common /legacy1/ a, b
integer sum
common /legacy2/ sum
a = 1
b = 2
sum = a + b
if ((sum) /= 3) then
    print *, "FAIL: want [3] got [", sum, "]"
    stop 1
end if
end program statement_f77_legacy_compat_common_blocks
