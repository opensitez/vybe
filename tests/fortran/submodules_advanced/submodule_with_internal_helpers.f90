! vybe-test: fortran/submodules_advanced/submodule_with_internal_helpers
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module stats_iface
    implicit none
    interface
        module function mean(a) result(m)
            real, intent(in) :: a(:)
            real :: m
        end function mean
        module function variance(a) result(v)
            real, intent(in) :: a(:)
            real :: v
        end function variance
    end interface
end module stats_iface

submodule (stats_iface) stats_impl
    implicit none
contains
    module function mean(a) result(m)
        real, intent(in) :: a(:)
        real :: m
        m = sum(a) / real(size(a))
    end function mean

    module function variance(a) result(v)
        real, intent(in) :: a(:)
        real :: v
        real :: m
        m = mean(a)
        v = sum((a - m)**2) / real(size(a))
    end function variance
end submodule stats_impl

program test
    use stats_iface
    real :: data(5) = [1.0, 2.0, 3.0, 4.0, 5.0]
    if ((int(mean(data))) /= 3) then
    print *, "FAIL: want [3] got [", int(mean(data)), "]"
    stop 1
end if
    if ((int(variance(data))) /= 2) then
    print *, "FAIL: want [2] got [", int(variance(data)), "]"
    stop 1
end if
end program test
