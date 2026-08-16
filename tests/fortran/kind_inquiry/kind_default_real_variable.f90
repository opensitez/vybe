! vybe-test: fortran/kind_inquiry/kind_default_real_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real :: x = 1.5
if ((kind(x)) /= 4) then
    print *, "FAIL: want [4] got [", kind(x), "]"
    stop 1
end if
end program t
