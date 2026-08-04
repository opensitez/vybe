! vybe-test: fortran/fortran2003_extended/compile_polymorphic_select_type_deferred_child
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

module poly
    implicit none
    type, abstract :: Op
    contains
        procedure(run_iface), deferred :: run
    end type Op
    abstract interface
        integer function run_iface(self) result(v)
            import Op
            class(Op), intent(in) :: self
        end function run_iface
    end interface
    type, extends(Op) :: Inc
        integer :: step = 1
    contains
        procedure :: run => inc_run
    end type Inc
contains
    function inc_run(self) result(v)
        class(Inc), intent(in) :: self
        v = self%step
    end function inc_run
end module poly

program t
    use poly
    class(Op), allocatable :: job
    allocate(Inc :: job)
    select type(job)
    type is (Inc)
        print *, job%run()
  class default
        print *, 0
    end select
end program t
