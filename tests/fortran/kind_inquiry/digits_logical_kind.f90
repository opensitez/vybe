! vybe-test: fortran/kind_inquiry/digits_logical_kind
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
logical :: l
if ((kind(l)) /= 8) then
    print *, "FAIL: want [8] got [", kind(l), "]"
    stop 1
end if
if ((kind(bit_size(l))) /= 8) then
    print *, "FAIL: want [8] got [", kind(bit_size(l)), "]"
    stop 1
end if
end program t
