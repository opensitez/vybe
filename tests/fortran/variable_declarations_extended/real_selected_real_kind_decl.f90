! vybe-test: fortran/variable_declarations_extended/real_selected_real_kind_decl
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: rk = selected_real_kind(6)
real(kind=rk) :: z = 0.125
if (abs((z) - 0.125) > 1.0e-6) then
    print *, "FAIL: want [0.125] got [", z, "]"
    stop 1
end if
end program t
