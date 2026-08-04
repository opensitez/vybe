! vybe-test: fortran/arrays/intrinsic_maxloc
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [3, 1, 9, 1, 5]
    integer :: loc(1)
    loc = maxloc(a)
    print *, loc(1)
end program test
