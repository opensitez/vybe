! vybe-test: fortran/strings_extended/char_compare_eq
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=5) :: a = 'hello'
character(len=5) :: b = 'hello'
if ((a == b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", a == b, "]"
    stop 1
end if
end program t
