! vybe-test: fortran/kind_inquiry/digits_int_kind_eight_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer(kind=8) :: x = 0_8
if ((digits(x)) /= 18) then
    print *, "FAIL: want [18] got [", digits(x), "]"
    stop 1
end if
end program t
