! vybe-test: fortran/fortran2008/block_allocatable
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    block
        integer, allocatable :: arr(:)
        allocate(arr(5))
        arr = [1, 2, 3, 4, 5]
        print *, sum(arr)
    end block
end program test
