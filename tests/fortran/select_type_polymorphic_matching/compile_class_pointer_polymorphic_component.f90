! vybe-test: fortran/select_type_polymorphic_matching/compile_class_pointer_polymorphic_component
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Base
        integer :: tag = 1
    end type Base
    type, extends(Base) :: Ext
        integer :: extra = 9
    end type Ext
    type :: Holder
        class(Base), pointer :: item => null()
    end type Holder
    type(Holder) :: h
    type(Ext), target :: e
    h%item => e
    print *, h%item%tag
end program t
