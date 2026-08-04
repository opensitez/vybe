! vybe-test: fortran/formatting/fmt_sign_12
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
integer::x=1
write(*,'(SP,I3)') x
end program p
