! vybe-test: fortran/control_flow/logical_operators
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
    logical :: a, b
    a = .true.
    b = .false.
    if (a .and. .not. b) then
        if (trim("yes") /= "yes") then
    print *, "FAIL: want [yes] got [", "yes", "]"
    stop 1
end if
    end if
end program test
