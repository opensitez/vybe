! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_integer_scaling_result_name
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_integer_scaling_result_name
    integer function scaled(v) result(output)
        integer, intent(in) :: v
        output = v * 3
    end function scaled
    if ((scaled(4)) /= 12) then
    print *, "FAIL: want [12] got [", scaled(4), "]"
    stop 1
end if
end program procedure_result_rename_patterns_integer_scaling_result_name
