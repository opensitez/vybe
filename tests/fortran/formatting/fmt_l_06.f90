! vybe-test: fortran/formatting/fmt_l_06
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
logical::x=.true.
write(*,'(L3)') x
end program p
