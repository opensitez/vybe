! vybe-test: fortran/strings_extended/string_slice_len_trim_chain
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=12) :: s
s = '  trim_me  '
if ((len_trim(s(3:))) /= 7) then
    print *, "FAIL: want [7] got [", len_trim(s(3:)), "]"
    stop 1
end if
if ((len_trim(s(:4))) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(s(:4)), "]"
    stop 1
end if
if ((len_trim(s(3:6))) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(s(3:6)), "]"
    stop 1
end if
end program t
