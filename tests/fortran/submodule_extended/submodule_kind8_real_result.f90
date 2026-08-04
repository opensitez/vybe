! vybe-test: fortran/submodule_extended/submodule_kind8_real_result
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module dbl_iface
    implicit none
    interface
        module function halve(x) result(r)
            real(kind=8), intent(in) :: x
            real(kind=8) :: r
        end function halve
    end interface
end module dbl_iface

submodule (dbl_iface) dbl_impl
contains
    module function halve(x) result(r)
        real(kind=8), intent(in) :: x
        real(kind=8) :: r
        r = x / 2.0d0
    end function halve
end submodule dbl_impl

program t
    use dbl_iface
    if ((int(halve(9.0d0))) /= 4) then
    print *, "FAIL: want [4] got [", int(halve(9.0d0)), "]"
    stop 1
end if
end program t
