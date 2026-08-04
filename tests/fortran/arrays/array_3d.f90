! vybe-test: fortran/arrays/array_3d
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: t(2,2,2)
    t(1,1,1) = 111
    print *, t(1,1,1)
end program test
