! vybe-test: fortran/kind_inquiry/bit_size_logical_literal
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((bit_size(.true.)) /= 32) then
    print *, "FAIL: want [32] got [", bit_size(.true.), "]"
    stop 1
end if
end program t
