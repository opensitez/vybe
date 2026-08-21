! vybe-test: fortran/move_alloc/move_alloc_basic
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    integer, allocatable :: a(:), b(:)
    allocate(a(3))
    a = [1, 2, 3]
    call move_alloc(a, b)
    print *, b(1)
    print *, allocated(a)
end program test
