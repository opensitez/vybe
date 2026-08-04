! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_nested_function_result_variable
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_nested_function_result_variable
    integer function outer(v) result(total)
        integer, intent(in) :: v
        integer :: helper
        helper = v + 1
        total = helper * 2
    end function outer
    if ((outer(3)) /= 8) then
    print *, "FAIL: want [8] got [", outer(3), "]"
    stop 1
end if
end program procedure_result_rename_patterns_nested_function_result_variable
