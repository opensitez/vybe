! vybe-test: fortran/strings_extended/shared_str_split_direct_index_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
if (trim(trim(str_split('alpha:beta:gamma', ':')(1))) /= "alpha") then
    print *, "FAIL: want [alpha] got [", trim(str_split('alpha:beta:gamma', ':')(1)), "]"
    stop 1
end if
if (trim(trim(str_split('alpha:beta:gamma', ':')(3))) /= "gamma") then
    print *, "FAIL: want [gamma] got [", trim(str_split('alpha:beta:gamma', ':')(3)), "]"
    stop 1
end if
end program t
