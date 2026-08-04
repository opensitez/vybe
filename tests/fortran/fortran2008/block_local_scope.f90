! vybe-test: fortran/fortran2008/block_local_scope
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: i = 10
    block
        integer :: i   ! shadows outer i
        i = 99
        print *, i
    end block
    print *, i  ! outer i unchanged
end program test
