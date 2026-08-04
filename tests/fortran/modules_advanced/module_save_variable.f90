! vybe-test: fortran/modules_advanced/module_save_variable
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module counter_mod
    implicit none
    integer, save :: count = 0
contains
    subroutine increment()
        count = count + 1
    end subroutine increment
    function get_count() result(c)
        integer :: c
        c = count
    end function get_count
end module counter_mod

program test
    use counter_mod
    call increment()
    call increment()
    call increment()
    print *, get_count()
end program test
