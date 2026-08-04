! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_procedure_argument_priority
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_procedure_argument_priority
    integer :: x
    x = 8
    if ((compute(x)) /= 10) then
    print *, "FAIL: want [10] got [", compute(x), "]"
    stop 1
end if
contains
    integer function compute(x)
        integer, intent(in) :: x
        integer :: local
        local = x + 2
        compute = local
    end function compute
end program variable_shadowing_resolution_rules_procedure_argument_priority
