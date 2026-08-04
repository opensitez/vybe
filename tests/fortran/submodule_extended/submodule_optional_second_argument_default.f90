! vybe-test: fortran/submodule_extended/submodule_optional_second_argument_default
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module opt_iface
    implicit none
    interface
        module function bump(n, step) result(r)
            integer, intent(in) :: n
            integer, optional, intent(in) :: step
            integer :: r
        end function bump
    end interface
end module opt_iface

submodule (opt_iface) opt_impl
contains
    module function bump(n, step) result(r)
        integer, intent(in) :: n
        integer, optional, intent(in) :: step
        integer :: r
        if (present(step)) then
            r = n + step
        else
            r = n + 1
        end if
    end function bump
end submodule opt_impl

program t
    use opt_iface
    if ((bump(5)) /= 6) then
    print *, "FAIL: want [6] got [", bump(5), "]"
    stop 1
end if
    if ((bump(5, 3)) /= 8) then
    print *, "FAIL: want [8] got [", bump(5, 3), "]"
    stop 1
end if
end program t
