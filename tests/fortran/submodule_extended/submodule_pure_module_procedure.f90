! vybe-test: fortran/submodule_extended/submodule_pure_module_procedure
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module pure_iface
    implicit none
    interface
        module function pure_add(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function pure_add
    end interface
end module pure_iface

submodule (pure_iface) pure_impl
    implicit none
contains
    module function pure_add(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a + b
    end function pure_add
end submodule pure_impl

program t
    use pure_iface
    if ((pure_add(3, 4)) /= 7) then
    print *, "FAIL: want [7] got [", pure_add(3, 4), "]"
    stop 1
end if
end program t
