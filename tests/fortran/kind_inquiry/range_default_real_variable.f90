! vybe-test: fortran/kind_inquiry/range_default_real_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real :: x = 0.0
if ((range(x)) /= 307) then
    print *, "FAIL: want [307] got [", range(x), "]"
    stop 1
end if
end program t
