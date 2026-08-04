! vybe-test: fortran/formatting/fmt_i_01
! origin: languages/fortran/tests/fortran/test_formatting.rs
program p
integer :: x=1
write(*,'(I3)') x
end program p
