! vybe-test: fortran/strings_extended/shared_array_join_direct_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
if (trim(trim(array_join(str_split('alpha:beta:gamma', ':'), ' | '))) /= "alpha | beta | gamma") then
    print *, "FAIL: want [alpha | beta | gamma] got [", trim(array_join(str_split('alpha:beta:gamma', ':'), ' | ')), "]"
    stop 1
end if
end program t
