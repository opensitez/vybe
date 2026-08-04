! vybe-test: fortran/submodule_extended/submodule_character_scalar_result
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module char_iface
    implicit none
    interface
        module function first_char(s) result(c)
            character(len=*), intent(in) :: s
            character(len=1) :: c
        end function first_char
    end interface
end module char_iface

submodule (char_iface) char_impl
contains
    module function first_char(s) result(c)
        character(len=*), intent(in) :: s
        character(len=1) :: c
        c = s(1:1)
    end function first_char
end submodule char_impl

program t
    use char_iface
    if (trim(first_char('delta')) /= "d") then
    print *, "FAIL: want [d] got [", first_char('delta'), "]"
    stop 1
end if
end program t
