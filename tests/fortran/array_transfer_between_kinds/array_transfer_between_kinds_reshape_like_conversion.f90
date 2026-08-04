! vybe-test: fortran/array_transfer_between_kinds/array_transfer_between_kinds_reshape_like_conversion
! origin: languages/fortran/tests/fortran/test_array_transfer_between_kinds.rs

program array_transfer_between_kinds_reshape_like_conversion
    integer :: a(2)
    integer :: b(2)
    a = (/11, 22/)
    b = transfer(a, b)
    if ((b(1)) /= 11) then
    print *, "FAIL: want [11] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 22) then
    print *, "FAIL: want [22] got [", b(2), "]"
    stop 1
end if
end program array_transfer_between_kinds_reshape_like_conversion
