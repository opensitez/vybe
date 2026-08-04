! vybe-test: fortran/strings_extended/index_from_back
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=20) :: s = 'abcabc'
if ((index(s, 'bc', .true.)) /= 5) then
    print *, "FAIL: want [5] got [", index(s, 'bc', .true.), "]"
    stop 1
end if
end program t
