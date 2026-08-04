! vybe-test: fortran/formatting/fmt_blank_13
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
integer::x=1
write(*,'(BN,I3)') x
end program p
