! vybe-test: fortran/fortran2003_extended/compile_deferred_two_abstract_procedures
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

module iface2
    implicit none
    type, abstract :: Expr
    contains
        procedure(eval_iface), deferred :: eval
        procedure(arity_iface), deferred :: arity
    end type Expr

    abstract interface
        integer function eval_iface(self) result(v)
            import Expr
            class(Expr), intent(in) :: self
        end function eval_iface
        integer function arity_iface(self) result(n)
            import Expr
            class(Expr), intent(in) :: self
        end function arity_iface
    end interface
end module iface2

program t
    use iface2
    print *, "ok"
end program t
