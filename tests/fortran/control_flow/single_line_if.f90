! vybe-test: fortran/control_flow/single_line_if
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
    integer :: x
    x = 5
    if (x > 3) print *, "big"
end program test
