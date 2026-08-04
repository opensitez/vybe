! vybe-test: fortran/bits_f2008/merge_bits_alternating
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs

program test
    integer :: result
    result = merge_bits(int(z'AAAA'), int(z'5555'), int(z'FF00'))
    print *, result
end program test
