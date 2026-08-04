! vybe-test: fortran/submodule_extended/submodule_logical_result_function
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module even_iface
    implicit none
    interface
        module function is_even(n) result(flag)
            integer, intent(in) :: n
            logical :: flag
        end function is_even
    end interface
end module even_iface

submodule (even_iface) even_impl
contains
    module function is_even(n) result(flag)
        integer, intent(in) :: n
        logical :: flag
        flag = mod(n, 2) == 0
    end function is_even
end submodule even_impl

program t
    use even_iface
    if ((is_even(8)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", is_even(8), "]"
    stop 1
end if
    if ((is_even(7)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", is_even(7), "]"
    stop 1
end if
end program t
