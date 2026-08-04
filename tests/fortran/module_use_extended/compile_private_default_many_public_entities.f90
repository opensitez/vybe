! vybe-test: fortran/module_use_extended/compile_private_default_many_public_entities
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module access_mix
    implicit none
    private
    public :: a, b, show
    integer :: a = 1
    integer :: b = 2
    integer :: c = 3
contains
    subroutine show()
        print *, a + b
    end subroutine show
end module access_mix

program t
    use access_mix
    call show()
end program t
