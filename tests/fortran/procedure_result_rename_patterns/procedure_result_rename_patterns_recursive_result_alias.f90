! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_recursive_result_alias
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_recursive_result_alias
    integer function fact(n) result(r)
        integer, intent(in) :: n
        if (n <= 1) then
            r = 1
        else
            r = n * fact(n - 1)
        end if
    end function fact
    if ((fact(4)) /= 24) then
    print *, "FAIL: want [24] got [", fact(4), "]"
    stop 1
end if
end program procedure_result_rename_patterns_recursive_result_alias
