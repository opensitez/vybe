! vybe-test: fortran/formatting/fmt_e_03
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
real :: x=1.5
write(*,'(E10.2)') x
end program p
