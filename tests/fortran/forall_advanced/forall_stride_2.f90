! vybe-test: fortran/forall_advanced/forall_stride_2
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(10) = 0
    forall (i = 1:10:2)
        a(i) = i
    end forall
    print *, a(1)
    print *, a(3)
    print *, a(2)
end program test
