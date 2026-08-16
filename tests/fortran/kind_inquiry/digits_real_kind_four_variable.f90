! vybe-test: fortran/kind_inquiry/digits_real_kind_four_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=4) :: x = 0.0_4
if ((digits(x)) /= 24) then
    print *, "FAIL: want [24] got [", digits(x), "]"
    stop 1
end if
end program t
