! vybe-test: fortran/kind_inquiry/digits_default_real_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real :: x = 0.0
if ((digits(x)) /= 24) then
    print *, "FAIL: want [24] got [", digits(x), "]"
    stop 1
end if
end program t
