! vybe-test: fortran/kind_inquiry/kind_real_kind_eight_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=8) :: x = 1.5_8
if ((kind(x)) /= 8) then
    print *, "FAIL: want [8] got [", kind(x), "]"
    stop 1
end if
end program t
