! vybe-test: fortran/name_resolution_renamed_imports/test_name_resolution_renamed_imports_prefers_renamed_symbol
! origin: languages/fortran/tests/fortran/test_name_resolution_renamed_imports.rs

module source_mod
    integer, parameter :: base = 14
end module

program test_name_resolution_renamed_imports
    use source_mod, only: exported => base
    if ((exported) /= 14) then
    print *, "FAIL: want [14] got [", exported, "]"
    stop 1
end if
end program test_name_resolution_renamed_imports
