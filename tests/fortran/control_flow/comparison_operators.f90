! vybe-test: fortran/control_flow/comparison_operators
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
    if (1 < 2) print *, "lt"
    if (2 > 1) print *, "gt"
    if (1 <= 1) print *, "le"
    if (1 >= 1) print *, "ge"
    if (1 == 1) print *, "eq"
    if (1 /= 2) print *, "ne"
end program test
