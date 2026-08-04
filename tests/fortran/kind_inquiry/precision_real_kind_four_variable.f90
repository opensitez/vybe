! vybe-test: fortran/kind_inquiry/precision_real_kind_four_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=4) :: x = 0.0_4
if ((precision(x)) /= 24) then
    print *, "FAIL: want [24] got [", precision(x), "]"
    stop 1
end if
end program t
