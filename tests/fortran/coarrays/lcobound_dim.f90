! vybe-test: fortran/coarrays/lcobound_dim
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[3:*]
    print *, lcobound(x)
end program test
