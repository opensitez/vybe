! vybe-test: fortran/abstract_types_deferred/deferred_binding
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

module iface_mod
    implicit none
    type, abstract :: Base
    contains
        procedure(greet_iface), deferred :: greet
    end type Base

    abstract interface
        subroutine greet_iface(self)
            import Base
            class(Base), intent(in) :: self
        end subroutine greet_iface
    end interface
end module iface_mod

program test
    print *, "ok"
end program test
