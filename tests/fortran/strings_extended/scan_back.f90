! vybe-test: fortran/strings_extended/scan_back
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hello'
if ((scan(s, 'aeiou', .true.)) /= 5) then
    print *, "FAIL: want [5] got [", scan(s, 'aeiou', .true.), "]"
    stop 1
end if
end program t
