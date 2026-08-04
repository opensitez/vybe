! vybe-test: fortran/variable_declarations_extended/real_kind_4_init
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
real(kind=4) :: x = 2.5
if (abs((x) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", x, "]"
    stop 1
end if
end program t
