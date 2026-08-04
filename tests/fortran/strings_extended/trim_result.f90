! vybe-test: fortran/strings_extended/trim_result
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hello'
if (trim(trim(s)) /= "hello") then
    print *, "FAIL: want [hello] got [", trim(s), "]"
    stop 1
end if
end program t
