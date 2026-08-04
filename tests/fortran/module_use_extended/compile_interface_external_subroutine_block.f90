! vybe-test: fortran/module_use_extended/compile_interface_external_subroutine_block
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module ext_iface
    implicit none
    interface
        subroutine external_log(msg)
            character(len=*), intent(in) :: msg
        end subroutine external_log
    end interface
contains
    subroutine relay(msg)
        character(len=*), intent(in) :: msg
        call external_log(msg)
    end subroutine relay
end module ext_iface

subroutine external_log(msg)
    character(len=*), intent(in) :: msg
    print *, trim(msg)
end subroutine external_log

program t
    use ext_iface
    call relay("hi")
end program t
