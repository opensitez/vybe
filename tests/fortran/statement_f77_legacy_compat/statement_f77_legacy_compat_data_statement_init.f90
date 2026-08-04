! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_data_statement_init
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_data_statement_init
integer i, j
data i /1/, j /2/
if ((i + j) /= 3) then
    print *, "FAIL: want [3] got [", i + j, "]"
    stop 1
end if
end program statement_f77_legacy_compat_data_statement_init
