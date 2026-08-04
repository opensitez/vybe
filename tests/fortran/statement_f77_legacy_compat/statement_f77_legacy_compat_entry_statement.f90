! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_entry_statement
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_entry_statement
integer :: sum
sum = value(1, 2) + value2(3, 4)
if ((sum) /= 15) then
    print *, "FAIL: want [15] got [", sum, "]"
    stop 1
end if
contains
integer function value(a, b)
integer a, b
value = a + b
return
entry value2(a, b)
value2 = a * b
end function value
end program statement_f77_legacy_compat_entry_statement
