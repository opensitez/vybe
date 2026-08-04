! vybe-test: fortran/kind_inquiry/kind_double_literal_one_point_zero_d0
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((kind(1.0d0)) /= 8) then
    print *, "FAIL: want [8] got [", kind(1.0d0), "]"
    stop 1
end if
end program t
