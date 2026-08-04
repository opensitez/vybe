! vybe-test: fortran/kinds/kind_real_array
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    real(kind=8) :: v(3) = [1.0_8, 2.0_8, 3.0_8]
    print *, v(1)
end program test
