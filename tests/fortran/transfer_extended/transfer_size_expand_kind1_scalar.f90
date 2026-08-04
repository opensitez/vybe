! vybe-test: fortran/transfer_extended/transfer_size_expand_kind1_scalar
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer(kind=1) :: source = 42_1
integer :: target(2)
target = transfer(source, target, 2)
if ((target(1)) /= 42) then
    print *, "FAIL: want [42] got [", target(1), "]"
    stop 1
end if
if ((target(2)) /= 0) then
    print *, "FAIL: want [0] got [", target(2), "]"
    stop 1
end if
end program t
