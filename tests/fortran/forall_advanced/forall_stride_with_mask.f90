! vybe-test: fortran/forall_advanced/forall_stride_with_mask
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(20) = 0
    forall (i = 1:20:2, i > 10)
        a(i) = i
    end forall
    print *, a(11)
    print *, a(9)
end program test
