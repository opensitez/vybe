! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_fixed_width_integers
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_fixed_width_integers
integer*2 short
integer*8 long
short = 2
long = 3
if ((short) /= 2) then
    print *, "FAIL: want [2] got [", short, "]"
    stop 1
end if
if ((long) /= 3) then
    print *, "FAIL: want [3] got [", long, "]"
    stop 1
end if
end program statement_f77_legacy_compat_fixed_width_integers
