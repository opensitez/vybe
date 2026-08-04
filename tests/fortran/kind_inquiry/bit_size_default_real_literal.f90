! vybe-test: fortran/kind_inquiry/bit_size_default_real_literal
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((bit_size(0.0)) /= 64) then
    print *, "FAIL: want [64] got [", bit_size(0.0), "]"
    stop 1
end if
end program t
