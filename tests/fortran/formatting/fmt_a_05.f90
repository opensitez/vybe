! vybe-test: fortran/formatting/fmt_a_05
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
character(len=3)::s='abc'
write(*,'(A)') s
end program p
