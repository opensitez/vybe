! vybe-test: fortran/kind_inquiry/precision_default_real_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real :: x = 0.0
if ((precision(x)) /= 6) then
    print *, "FAIL: want [6] got [", precision(x), "]"
    stop 1
end if
end program t
