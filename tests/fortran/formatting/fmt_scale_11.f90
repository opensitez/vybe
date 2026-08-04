! vybe-test: fortran/formatting/fmt_scale_11
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
real::x=1.23
write(*,'(1P,E10.2)') x
end program p
