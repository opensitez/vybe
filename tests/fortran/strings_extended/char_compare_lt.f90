! vybe-test: fortran/strings_extended/char_compare_lt
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=5) :: a = 'apple'
character(len=6) :: b = 'banana'
if ((a < b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", a < b, "]"
    stop 1
end if
end program t
