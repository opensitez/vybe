! vybe-test: fortran/pointer_alloc_extended/compile_allocatable_function_result
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

function doubled(n) result(out)
    integer, intent(in) :: n
    integer, allocatable :: out(:)
    allocate(out(n))
    out = [(2 * i, i = 1, n)]
end function doubled

program t
    integer, allocatable :: v(:)
    v = doubled(3)
    print *, v(3)
end program t
