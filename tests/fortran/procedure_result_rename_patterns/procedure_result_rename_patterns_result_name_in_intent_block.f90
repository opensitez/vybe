! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_result_name_in_intent_block
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_result_name_in_intent_block
    integer function accumulate(a, b) result(total)
        integer, intent(in) :: a
        integer, intent(in) :: b
        total = a + b
        if (a > b) total = total + 1
    end function accumulate
    if ((accumulate(2, 5)) /= 7) then
    print *, "FAIL: want [7] got [", accumulate(2, 5), "]"
    stop 1
end if
    if ((accumulate(9, 1)) /= 11) then
    print *, "FAIL: want [11] got [", accumulate(9, 1), "]"
    stop 1
end if
end program procedure_result_rename_patterns_result_name_in_intent_block
