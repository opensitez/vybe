! vybe-test: fortran/block_construct_extended/block_nested
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: x = 1
    block
        integer :: y
        y = x + 10
        block
            integer :: z
            z = y + 100
            print *, z
        end block
    end block
end program test
