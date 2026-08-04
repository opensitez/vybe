! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_name_hides_intrinsic
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_name_hides_intrinsic
    integer :: sum
    sum = 5
    if ((bump(sum)) /= 6) then
    print *, "FAIL: want [6] got [", bump(sum), "]"
    stop 1
end if
contains
    integer function bump(sum)
        integer, intent(in) :: sum
        bump = sum + 1
    end function bump
end program variable_shadowing_resolution_rules_name_hides_intrinsic
