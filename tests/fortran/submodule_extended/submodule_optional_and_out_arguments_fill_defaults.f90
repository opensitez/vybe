! vybe-test: fortran/submodule_extended/submodule_optional_and_out_arguments_fill_defaults
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module out_iface
    implicit none
    interface
        module subroutine fill_pair(base, a, b, result_out)
            integer, intent(in) :: base
            integer, optional, intent(in) :: a, b
            integer, intent(out) :: result_out
        end subroutine fill_pair
    end interface
end module out_iface

submodule (out_iface) out_impl
contains
    module subroutine fill_pair(base, a, b, result_out)
        integer, intent(in) :: base
        integer, optional, intent(in) :: a, b
        integer, intent(out) :: result_out
        if (present(a) .and. present(b)) then
            result_out = base + a + b
        else
            result_out = base + 1
        end if
    end subroutine fill_pair
end submodule out_impl

program t
    use out_iface
    integer :: x
    call fill_pair(6, result_out=x)
    if ((x) /= 7) then
    print *, "FAIL: want [7] got [", x, "]"
    stop 1
end if
    call fill_pair(6, a=2, b=3, result_out=x)
    if ((x) /= 11) then
    print *, "FAIL: want [11] got [", x, "]"
    stop 1
end if
end program t
