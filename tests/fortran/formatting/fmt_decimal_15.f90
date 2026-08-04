! vybe-test: fortran/formatting/fmt_decimal_15
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
real::x=1.2
write(*,'(F5.1,DECIMAL="POINT")') x
end program p
