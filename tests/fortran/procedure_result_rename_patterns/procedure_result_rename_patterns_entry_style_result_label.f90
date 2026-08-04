! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_entry_style_result_label
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_entry_style_result_label
    integer function normalize(v) result(value)
        integer, intent(in) :: v
        value = v
        if (v < 0) value = -v
    end function normalize
    if ((normalize(-12)) /= 12) then
    print *, "FAIL: want [12] got [", normalize(-12), "]"
    stop 1
end if
    if ((normalize(15)) /= 15) then
    print *, "FAIL: want [15] got [", normalize(15), "]"
    stop 1
end if
end program procedure_result_rename_patterns_entry_style_result_label
