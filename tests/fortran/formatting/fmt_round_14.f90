! vybe-test: fortran/formatting/fmt_round_14
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
real::x=1.2
write(*,'(F5.1,ROUND="UP")') x
end program p
