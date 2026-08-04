! vybe-test: fortran/module_use_extended/compile_generic_interface_three_procedures
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module g3
    implicit none
    interface pick
        module procedure pick_int, pick_real, pick_logical
    end interface
contains
    function pick_int(v) result(r)
        integer, intent(in) :: v
        integer :: r
        r = v
    end function pick_int
    function pick_real(v) result(r)
        real, intent(in) :: v
        real :: r
        r = v
    end function pick_real
    function pick_logical(v) result(r)
        logical, intent(in) :: v
        logical :: r
        r = v
    end function pick_logical
end module g3

program t
    use g3
    print *, pick(1)
    print *, int(pick(2.0))
    print *, pick(.true.)
end program t
