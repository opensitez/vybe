! vybe-test: fortran/strings_extended/str_split_multi_char_delimiter_chain
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=20), allocatable :: parts(:)
parts = str_split('x--y--z--', '--')
if ((size(parts)) /= 4) then
    print *, "FAIL: want [4] got [", size(parts), "]"
    stop 1
end if
if (trim(trim(parts(1))) /= "x") then
    print *, "FAIL: want [x] got [", trim(parts(1)), "]"
    stop 1
end if
if (trim(trim(parts(2))) /= "y") then
    print *, "FAIL: want [y] got [", trim(parts(2)), "]"
    stop 1
end if
if ((len_trim(trim(parts(4)))) /= 0) then
    print *, "FAIL: want [0] got [", len_trim(trim(parts(4))), "]"
    stop 1
end if
end program t
