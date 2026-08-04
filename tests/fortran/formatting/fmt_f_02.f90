! vybe-test: fortran/formatting/fmt_f_02
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
real :: x=1.5
write(*,'(F5.1)') x
end program p
