! vybe-test: fortran/coarrays/lcobound_ucobound
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[2:4, *]
    print *, lcobound(x, 1)
    print *, ucobound(x, 1)
end program test
