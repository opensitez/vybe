! vybe-test: fortran/submodule_extended/submodule_reads_parent_module_variable
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module mult_iface
    implicit none
    integer :: factor = 7
    interface
        module function times_factor(n) result(r)
            integer, intent(in) :: n
            integer :: r
        end function times_factor
    end interface
end module mult_iface

submodule (mult_iface) mult_impl
contains
    module function times_factor(n) result(r)
        integer, intent(in) :: n
        integer :: r
        r = n * factor
    end function times_factor
end submodule mult_impl

program t
    use mult_iface
    if ((times_factor(6)) /= 42) then
    print *, "FAIL: want [42] got [", times_factor(6), "]"
    stop 1
end if
end program t
