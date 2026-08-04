! vybe-test: fortran/kinds/selected_real_kind_15
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: k = selected_real_kind(15, 307)
    real(kind=k) :: x = 1.23456789012345_k
    print *, x
end program test
