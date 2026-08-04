! vybe-test: fortran/pointer_alloc_extended/compile_nullify_reassociate_multiple_pointers
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

program t
    integer, target :: s = 1, t = 2, u = 3
    integer, pointer :: p => null(), q => null(), r => null()
    p => s
    q => t
    r => u
    nullify(p, q, r)
    p => t
    r => s
    print *, p
    print *, r
    print *, associated(q)
end program t
