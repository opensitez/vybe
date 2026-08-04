! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_interface_shadows
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_interface_shadows
    integer :: limit
    limit = 3
    if ((wrap(limit)) /= 4) then
    print *, "FAIL: want [4] got [", wrap(limit), "]"
    stop 1
end if
contains
    integer function wrap(value)
        integer, intent(in) :: value
        integer :: limit
        limit = value + 1
        wrap = limit
    end function wrap
end program variable_shadowing_resolution_rules_interface_shadows
