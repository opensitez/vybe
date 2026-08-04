! vybe-test: fortran/kind_inquiry/precision_real_kind_four_less_than_default
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=4) :: s = 0.0_4
real :: d = 0.0
if ((precision(s) < precision(d)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", precision(s) < precision(d), "]"
    stop 1
end if
end program t
