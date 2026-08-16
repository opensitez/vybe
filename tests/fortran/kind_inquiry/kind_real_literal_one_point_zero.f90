! vybe-test: fortran/kind_inquiry/kind_real_literal_one_point_zero
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((kind(1.0)) /= 4) then
    print *, "FAIL: want [4] got [", kind(1.0), "]"
    stop 1
end if
end program t
