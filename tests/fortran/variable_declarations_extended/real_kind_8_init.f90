! vybe-test: fortran/variable_declarations_extended/real_kind_8_init
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
real(kind=8) :: d = 1.25_8
if (abs((d) - 1.25) > 1.0e-6) then
    print *, "FAIL: want [1.25] got [", d, "]"
    stop 1
end if
end program t
