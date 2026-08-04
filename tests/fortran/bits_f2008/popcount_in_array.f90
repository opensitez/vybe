! vybe-test: fortran/bits_f2008/popcount_in_array
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: a(4) = [0, 1, 3, 7]
    integer :: i
    do i = 1, 4
        print *, popcount(a(i))
    end do
end program test
