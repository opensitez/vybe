! vybe-test: fortran/arrays/forall_basic
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5)
    forall (i = 1:5)
        a(i) = i * i
    end forall
    print *, a(3)
end program test
