! vybe-test: fortran/bits_f2008/leadz_power_of_two
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: i
    do i = 0, 7
        print *, leadz(2**i)
    end do
end program test
