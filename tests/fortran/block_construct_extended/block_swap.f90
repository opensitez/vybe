! vybe-test: fortran/block_construct_extended/block_swap
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a = 3, b = 7
    block
        integer :: tmp
        tmp = a
        a = b
        b = tmp
    end block
    print *, a
    print *, b
end program test
