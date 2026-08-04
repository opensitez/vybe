! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_data_repeat_syntax
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_data_repeat_syntax
integer i, j(3)
data i /3*2/
data (j(k), k=1,3) /1,2,3/
if ((i) /= 6) then
    print *, "FAIL: want [6] got [", i, "]"
    stop 1
end if
if ((j(1)) /= 1) then
    print *, "FAIL: want [1] got [", j(1), "]"
    stop 1
end if
if ((j(2)) /= 2) then
    print *, "FAIL: want [2] got [", j(2), "]"
    stop 1
end if
if ((j(3)) /= 3) then
    print *, "FAIL: want [3] got [", j(3), "]"
    stop 1
end if
end program statement_f77_legacy_compat_data_repeat_syntax
