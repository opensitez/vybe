! vybe-test: fortran/transfer/transfer_in_function
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    if (.not. (real_bits_roundtrip(1.0))) then
    print *, "FAIL: want [1] got [", real_bits_roundtrip(1.0), "]"
    stop 1
end if
contains
    logical function real_bits_roundtrip(x)
        real, intent(in) :: x
        integer :: n
        n = transfer(x, 0)
        real_bits_roundtrip = (transfer(n, 0.0) == x)
    end function real_bits_roundtrip
end program test
