! vybe-test: fortran/attributes/attr_pointer_runtime_aliasing
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_pointer_runtime_aliasing
    integer, target :: storage
    integer, pointer :: p
    storage = 33
    p => storage
    if ((p) /= 33) then
    print *, "FAIL: want [33] got [", p, "]"
    stop 1
end if
    p = 44
    if ((storage) /= 44) then
    print *, "FAIL: want [44] got [", storage, "]"
    stop 1
end if
end program attr_pointer_runtime_aliasing
