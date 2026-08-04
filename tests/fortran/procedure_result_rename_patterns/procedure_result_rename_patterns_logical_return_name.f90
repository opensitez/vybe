! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_logical_return_name
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_logical_return_name
    logical function has_even(v) result(flag)
        integer, intent(in) :: v
        flag = mod(v, 2) == 0
    end function has_even
    if ((has_even(8)) .neqv. .true.) then
    print *, "FAIL: want [True] got [", has_even(8), "]"
    stop 1
end if
    if ((has_even(9)) .neqv. .false.) then
    print *, "FAIL: want [False] got [", has_even(9), "]"
    stop 1
end if
end program procedure_result_rename_patterns_logical_return_name
