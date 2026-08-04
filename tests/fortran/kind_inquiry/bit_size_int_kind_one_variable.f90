! vybe-test: fortran/kind_inquiry/bit_size_int_kind_one_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer(kind=1) :: x = 0_1
if ((bit_size(x)) /= 8) then
    print *, "FAIL: want [8] got [", bit_size(x), "]"
    stop 1
end if
end program t
