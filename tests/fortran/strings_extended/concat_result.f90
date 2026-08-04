! vybe-test: fortran/strings_extended/concat_result
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: a = 'Hello'
character(len=10) :: b = ' World'
character(len=20) :: c
c = trim(a) // trim(b)
if (trim(trim(c)) /= "Hello World") then
    print *, "FAIL: want [Hello World] got [", trim(c), "]"
    stop 1
end if
end program t
