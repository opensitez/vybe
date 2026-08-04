! vybe-test: fortran/procedure_result_rename_patterns/procedure_result_rename_patterns_character_build_result_name
! origin: languages/fortran/tests/fortran/test_procedure_result_rename_patterns.rs

program procedure_result_rename_patterns_character_build_result_name
    character(len=16) function with_suffix(base) result(out)
        character(len=*), intent(in) :: base
        out = trim(base) // '_ok'
    end function with_suffix
    if (trim(trim(with_suffix('test'))) /= "test_ok") then
    print *, "FAIL: want [test_ok] got [", trim(with_suffix('test')), "]"
    stop 1
end if
end program procedure_result_rename_patterns_character_build_result_name
