! vybe-test: fortran/control_flow/single_line_if_false_path
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
    integer :: x
    x = 2
    if (x > 3) print *, "big"
    if (trim("after") /= "after") then
    print *, "FAIL: want [after] got [", "after", "]"
    stop 1
end if
end program test
