! vybe-test: fortran/transfer/transfer_in_subroutine
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    real :: x = 3.14
    logical :: same
    call get_bits(x, same)
    if ((same) /= 1) then
    print *, "FAIL: want [1] got [", same, "]"
    stop 1
end if
contains
    subroutine get_bits(x, same)
        real, intent(in) :: x
        logical, intent(out) :: same
        integer :: n
        n = transfer(x, 0)
        same = transfer(n, 0.0) == x
    end subroutine get_bits
end program test
