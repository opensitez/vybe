! vybe-test: fortran/kind_inquiry/bit_size_int_kind_two_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer(kind=2) :: x = 0_2
if ((bit_size(x)) /= 16) then
    print *, "FAIL: want [16] got [", bit_size(x), "]"
    stop 1
end if
end program t
