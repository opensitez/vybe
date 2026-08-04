! vybe-test: fortran/kind_inquiry/bit_size_real_kind_four_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=4) :: x = 0.0_4
if ((bit_size(x)) /= 32) then
    print *, "FAIL: want [32] got [", bit_size(x), "]"
    stop 1
end if
end program t
