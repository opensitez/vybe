! vybe-test: fortran/derived_type_oop_extended/compile_final_on_extended_child_type
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs

program t
    type :: Base
        integer :: id = 0
    contains
        final :: base_done
    end type Base
    type, extends(Base) :: Child
    contains
        final :: child_done
    end type Child
    type(Child) :: c
    c%id = 7
    print *, c%id
contains
    subroutine base_done(self)
        type(Base), intent(inout) :: self
        self%id = 0
    end subroutine base_done
    subroutine child_done(self)
        type(Child), intent(inout) :: self
        self%id = -1
    end subroutine child_done
end program t
