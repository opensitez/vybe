! vybe-test: fortran/module_use_extended/compile_public_private_mixed_module_entities
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module mixed_access
    implicit none
    private
    public :: visible_val, reveal
    integer :: hidden = 0
    integer, public :: visible_val = 4
contains
    function reveal() result(v)
        integer :: v
        v = hidden + visible_val
    end function reveal
end module mixed_access

program t
    use mixed_access
    print *, reveal()
end program t
