! vybe-test: fortran/strings_extended/shared_str_split_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=256), allocatable :: tokens(:)
tokens = str_split('alpha:beta:gamma', ':')
if ((size(tokens)) /= 3) then
    print *, "FAIL: want [3] got [", size(tokens), "]"
    stop 1
end if
if (trim(trim(tokens(1))) /= "alpha") then
    print *, "FAIL: want [alpha] got [", trim(tokens(1)), "]"
    stop 1
end if
if (trim(trim(tokens(3))) /= "gamma") then
    print *, "FAIL: want [gamma] got [", trim(tokens(3)), "]"
    stop 1
end if
end program t
