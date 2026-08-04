! vybe-test: fortran/module_use_extended/compile_module_interface_procedure_signature
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module sig_iface
    implicit none
    interface
        module function signed_add(a, b) result(r)
            integer, intent(in) :: a, b
            integer :: r
        end function signed_add
    end interface
end module sig_iface

program t
    print *, "ok"
end program t
