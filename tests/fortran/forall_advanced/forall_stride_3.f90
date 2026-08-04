! vybe-test: fortran/forall_advanced/forall_stride_3
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(12) = 0
    forall (i = 1:12:3)
        a(i) = i
    end forall
    print *, a(1)
    print *, a(4)
    print *, a(2)
end program test
