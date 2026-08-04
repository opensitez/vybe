! vybe-test: fortran/module_use_resolution/test_module_use_resolution_pulls_public_binding
! origin: languages/fortran/tests/fortran/test_module_use_resolution.rs

module math_constants
    integer, parameter :: magic = 11
end module

program test_module_use_resolution
    use math_constants
    if ((magic) /= 11) then
    print *, "FAIL: want [11] got [", magic, "]"
    stop 1
end if
end program test_module_use_resolution
