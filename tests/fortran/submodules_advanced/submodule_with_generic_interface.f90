! vybe-test: fortran/submodules_advanced/submodule_with_generic_interface
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module generic_iface
    implicit none
    interface norm
        module function norm_real(a) result(r)
            real, intent(in) :: a(:)
            real :: r
        end function norm_real
        module function norm_dbl(a) result(r)
            real(kind=8), intent(in) :: a(:)
            real(kind=8) :: r
        end function norm_dbl
    end interface norm
end module generic_iface

submodule (generic_iface) generic_impl
    implicit none
contains
    module function norm_real(a) result(r)
        real, intent(in) :: a(:)
        real :: r
        r = sqrt(sum(a**2))
    end function norm_real

    module function norm_dbl(a) result(r)
        real(kind=8), intent(in) :: a(:)
        real(kind=8) :: r
        r = sqrt(sum(a**2))
    end function norm_dbl
end submodule generic_impl

program test
    use generic_iface
    real :: v(3) = [3.0, 4.0, 0.0]
    if ((int(norm(v))) /= 5) then
    print *, "FAIL: want [5] got [", int(norm(v)), "]"
    stop 1
end if
end program test
