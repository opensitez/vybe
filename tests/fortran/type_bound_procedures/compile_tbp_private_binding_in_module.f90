! vybe-test: fortran/type_bound_procedures/compile_tbp_private_binding_in_module
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

module secrets
    implicit none
    type :: Vault
        integer :: code = 0
    contains
        procedure, private :: seal
        procedure :: open => unlock
    end type Vault
contains
    subroutine seal(self, c)
        class(Vault), intent(inout) :: self
        integer, intent(in) :: c
        self%code = c
    end subroutine seal
    subroutine unlock(self)
        class(Vault), intent(inout) :: self
        self%code = 0
    end subroutine unlock
end module secrets

program t
    use secrets
    type(Vault) :: v
    call v%open()
    print *, v%code
end program t
