! vybe-test: fortran/kind_inquiry/precision_real_kind_eight_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=8) :: x = 0.0_8
if ((precision(x)) /= 15) then
    print *, "FAIL: want [15] got [", precision(x), "]"
    stop 1
end if
end program t
