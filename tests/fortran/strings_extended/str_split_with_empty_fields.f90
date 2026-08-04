! vybe-test: fortran/strings_extended/str_split_with_empty_fields
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10), allocatable :: tokens(:)
tokens = str_split('a,,b', ',')
if ((size(tokens)) /= 3) then
    print *, "FAIL: want [3] got [", size(tokens), "]"
    stop 1
end if
if (trim(trim(tokens(1))) /= "a") then
    print *, "FAIL: want [a] got [", trim(tokens(1)), "]"
    stop 1
end if
if ((len_trim(trim(tokens(2)))) /= 0) then
    print *, "FAIL: want [0] got [", len_trim(trim(tokens(2))), "]"
    stop 1
end if
if (trim(trim(tokens(3))) /= "b") then
    print *, "FAIL: want [b] got [", trim(tokens(3)), "]"
    stop 1
end if
end program t
