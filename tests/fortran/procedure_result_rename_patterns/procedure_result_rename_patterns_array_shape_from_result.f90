! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_array_shape_from_result
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_array_shape_from_result
    integer function pick(i) result(v)
        integer, intent(in) :: i
        if (i < 0) v = 0
        if (i >= 0 .and. i <= 2) v = i * 2
        if (i > 2) v = 99
    end function pick
    if ((pick(-1)) /= 0) then
    print *, "FAIL: want [0] got [", pick(-1), "]"
    stop 1
end if
    if ((pick(1)) /= 2) then
    print *, "FAIL: want [2] got [", pick(1), "]"
    stop 1
end if
    if ((pick(4)) /= 99) then
    print *, "FAIL: want [99] got [", pick(4), "]"
    stop 1
end if
end program procedure_result_rename_patterns_array_shape_from_result
