! vybe-test: fortran/transfer/transfer_sequence_type_runtime_roundtrip
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    type, sequence :: RGB
        integer(kind=1) :: r, g, b, a
    end type RGB
    type(RGB) :: in_colour
    type(RGB) :: out_colour
    integer(kind=1) :: bytes(4)

    in_colour%r = 1_1
    in_colour%g = 2_1
    in_colour%b = 3_1
    in_colour%a = 4_1
    bytes = transfer(in_colour, bytes)
    out_colour = transfer(bytes, out_colour)

    if ((in_colour%r == out_colour%r) /= 1) then
    print *, "FAIL: want [1] got [", in_colour%r == out_colour%r, "]"
    stop 1
end if
    if ((in_colour%g == out_colour%g) /= 1) then
    print *, "FAIL: want [1] got [", in_colour%g == out_colour%g, "]"
    stop 1
end if
    if ((in_colour%b == out_colour%b) /= 1) then
    print *, "FAIL: want [1] got [", in_colour%b == out_colour%b, "]"
    stop 1
end if
    if ((in_colour%a == out_colour%a) /= 1) then
    print *, "FAIL: want [1] got [", in_colour%a == out_colour%a, "]"
    stop 1
end if
end program test
