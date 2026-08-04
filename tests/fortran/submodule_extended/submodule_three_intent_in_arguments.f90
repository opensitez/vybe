! vybe-test: fortran/submodule_extended/submodule_three_intent_in_arguments
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module sum3_iface
    implicit none
    interface
        module function sum3(a, b, c) result(s)
            integer, intent(in) :: a, b, c
            integer :: s
        end function sum3
    end interface
end module sum3_iface

submodule (sum3_iface) sum3_impl
contains
    module function sum3(a, b, c) result(s)
        integer, intent(in) :: a, b, c
        integer :: s
        s = a + b + c
    end function sum3
end submodule sum3_impl

program t
    use sum3_iface
    if ((sum3(2, 3, 4)) /= 9) then
    print *, "FAIL: want [9] got [", sum3(2, 3, 4), "]"
    stop 1
end if
end program t
