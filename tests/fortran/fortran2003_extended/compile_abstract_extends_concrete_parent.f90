! vybe-test: fortran/fortran2003_extended/compile_abstract_extends_concrete_parent
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

module hier
    implicit none
    type :: Entity
        integer :: uid = 0
    end type Entity
    type, abstract, extends(Entity) :: Drawable
    contains
        procedure(draw_iface), deferred :: draw
    end type Drawable

    abstract interface
        subroutine draw_iface(self)
            import Drawable
            class(Drawable), intent(in) :: self
        end subroutine draw_iface
    end interface
end module hier

program t
    use hier
    type(Entity) :: e
    e%uid = 7
    print *, e%uid
end program t
