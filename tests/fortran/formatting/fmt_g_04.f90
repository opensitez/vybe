! vybe-test: fortran/formatting/fmt_g_04
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
real :: x=1.5
write(*,'(G10.2)') x
end program p
