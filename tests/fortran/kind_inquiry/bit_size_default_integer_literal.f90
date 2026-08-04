! vybe-test: fortran/kind_inquiry/bit_size_default_integer_literal
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((bit_size(0)) /= 32) then
    print *, "FAIL: want [32] got [", bit_size(0), "]"
    stop 1
end if
end program t
