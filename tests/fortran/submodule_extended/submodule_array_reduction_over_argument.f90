! vybe-test: fortran/submodule_extended/submodule_array_reduction_over_argument
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module peak_iface
    implicit none
    interface
        module function peak(v) result(m)
            integer, intent(in) :: v(:)
            integer :: m
        end function peak
    end interface
end module peak_iface

submodule (peak_iface) peak_impl
contains
    module function peak(v) result(m)
        integer, intent(in) :: v(:)
        integer :: m
        m = v(1)
        if (v(2) > m) m = v(2)
        if (v(3) > m) m = v(3)
    end function peak
end submodule peak_impl

program t
    use peak_iface
    integer :: data(3)
    data = [3, 11, 7]
    if ((peak(data)) /= 11) then
    print *, "FAIL: want [11] got [", peak(data), "]"
    stop 1
end if
end program t
