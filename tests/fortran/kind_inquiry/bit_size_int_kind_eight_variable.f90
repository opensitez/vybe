! vybe-test: fortran/kind_inquiry/bit_size_int_kind_eight_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer(kind=8) :: x = 0_8
if ((bit_size(x)) /= 64) then
    print *, "FAIL: want [64] got [", bit_size(x), "]"
    stop 1
end if
end program t
