! vybe-test: fortran/fortran2008/block_basic
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: x = 5
    block
        integer :: temp
        temp = x * 2
        print *, temp
    end block
end program test
