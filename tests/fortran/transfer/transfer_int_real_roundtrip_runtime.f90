! vybe-test: fortran/transfer/transfer_int_real_roundtrip_runtime
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: original
    integer :: recovered
    real :: bits
    original = 42
    bits = transfer(original, 0.0)
    recovered = transfer(bits, 0)
    if ((recovered) /= 42) then
    print *, "FAIL: want [42] got [", recovered, "]"
    stop 1
end if
    if ((original == recovered) /= 1) then
    print *, "FAIL: want [1] got [", original == recovered, "]"
    stop 1
end if
end program test
