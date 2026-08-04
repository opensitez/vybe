! vybe-test: fortran/transfer_extended/transfer_array_byte_from_integer_scalar
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: n = 305419896
integer(kind=1) :: b(4)
b = transfer(n, b)
if ((int(b(1))) /= 120) then
    print *, "FAIL: want [120] got [", int(b(1)), "]"
    stop 1
end if
if ((int(b(2))) /= 86) then
    print *, "FAIL: want [86] got [", int(b(2)), "]"
    stop 1
end if
if ((int(b(3))) /= 52) then
    print *, "FAIL: want [52] got [", int(b(3)), "]"
    stop 1
end if
if ((int(b(4))) /= 18) then
    print *, "FAIL: want [18] got [", int(b(4)), "]"
    stop 1
end if
end program t
