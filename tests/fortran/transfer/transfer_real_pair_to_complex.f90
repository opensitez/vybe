! vybe-test: fortran/transfer/transfer_real_pair_to_complex
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    real :: pair(2) = [1.0, 2.0]
    complex :: c
    c = transfer(pair, c)
    if ((real(c)) /= 1) then
    print *, "FAIL: want [1] got [", real(c), "]"
    stop 1
end if
    if ((imag(c)) /= 2) then
    print *, "FAIL: want [2] got [", imag(c), "]"
    stop 1
end if
end program test
