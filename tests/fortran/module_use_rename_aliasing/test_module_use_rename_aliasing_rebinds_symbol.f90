! vybe-test: fortran/module_use_rename_aliasing/test_module_use_rename_aliasing_rebinds_symbol
! origin: languages/fortran/tests/fortran/test_module_use_rename_aliasing.rs

module alias_mod
    integer, parameter :: original = 21
end module

program test_module_use_rename_aliasing
    use alias_mod, only: short => original
    if ((short) /= 21) then
    print *, "FAIL: want [21] got [", short, "]"
    stop 1
end if
end program test_module_use_rename_aliasing
