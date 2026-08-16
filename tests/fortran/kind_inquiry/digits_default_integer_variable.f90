! vybe-test: fortran/kind_inquiry/digits_default_integer_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: x = 0
if ((digits(x)) /= 31) then
    print *, "FAIL: want [31] got [", digits(x), "]"
    stop 1
end if
end program t
