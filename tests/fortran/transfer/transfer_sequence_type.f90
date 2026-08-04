! vybe-test: fortran/transfer/transfer_sequence_type
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    type, sequence :: RGB
        integer(kind=1) :: r, g, b, a
    end type RGB
    type(RGB) :: color
    integer :: packed
    color%r = 255_1; color%g = 0_1; color%b = 128_1; color%a = 255_1
    packed = transfer(color, 0)
    if ((color%r) /= 255) then
    print *, "FAIL: want [255] got [", color%r, "]"
    stop 1
end if
    if ((color%g) /= 0) then
    print *, "FAIL: want [0] got [", color%g, "]"
    stop 1
end if
    if ((color%b) /= 128) then
    print *, "FAIL: want [128] got [", color%b, "]"
    stop 1
end if
    if ((color%a) /= 255) then
    print *, "FAIL: want [255] got [", color%a, "]"
    stop 1
end if
end program test
