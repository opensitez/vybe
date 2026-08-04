! vybe-test: fortran/module_use_extended/compile_module_interface_block_standalone
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module standalone_iface
    implicit none
    interface
        function area_circle(r) result(a)
            real, intent(in) :: r
            real :: a
        end function area_circle
    end interface
contains
    function scaled_area(r, s) result(a)
        real, intent(in) :: r, s
        real :: a
        a = area_circle(r) * s
    end function scaled_area
end module standalone_iface

function area_circle(r) result(a)
    real, intent(in) :: r
    real :: a
    a = 3.0 * r * r
end function area_circle

program t
    use standalone_iface
    print *, int(scaled_area(2.0, 1.0))
end program t
