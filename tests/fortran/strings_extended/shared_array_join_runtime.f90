! vybe-test: fortran/strings_extended/shared_array_join_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=256), allocatable :: tokens(:)
tokens = str_split('alpha:beta:gamma', ':')
if (trim(trim(array_join(tokens, ' | '))) /= "alpha | beta | gamma") then
    print *, "FAIL: want [alpha | beta | gamma] got [", trim(array_join(tokens, ' | ')), "]"
    stop 1
end if
end program t
